from pyral import Rally, rallySettings, rallyWorkset
import sys
import csv

"""
Given a US find the parents
US > TF > FEA > PRJ > PGM
stop at PGM or null

"""

errout = sys.stderr.write

gFields = [ #"FormattedID",
        "Name", 
        "Project.Name",
        "PlanEstimate",
        "ScheduleState",
        "Iteration.Name",
        "LeafStoryCount",
        "AcceptedLeafStoryCount",
        "UnEstimatedLeafStoryCount",
        "LeafStoryPlanEstimateTotal",
        "AcceptedLeafStoryPlanEstimateTotal",
        #"State",
        #"TTSDefectState",
        ]

def returnAttrib(item, attr, default=""):
    locAttr = attr.split('.')
    if len(locAttr) == 1:
        return getattr(item, locAttr[0], default)
    else:
        return getattr(getattr(item, locAttr[0], ""), locAttr[1], default)

def printHelp():
    print ("USAGE: program <Search Token>")
    print ("    -h    Help")
            

def main(args):
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]
    if len(args) < 1:
        errout('ERROR: Wrong number of arguments\n')
        printHelp()
        sys.exit(3)

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

    bPrintStatus = True
    fileName = 'out.csv'

    with open (fileName, 'w', newline='', encoding="utf-8") as csvfile:
        #write header
        outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        outrow = ["Type", "PGM", "PRJ", "FEA.ID", "TF.ID", "US.ID"] + \
                    [field for field in gFields]
        outfile.writerow(outrow)

        for token in args:
            processQueryTokens(rally, bPrintStatus, outfile, token)



def processQueryTokens(rally, bPrintStatus, outfile, token):
    queryToken = token.upper()  #only search first one for now
    if queryToken[0] == "D":
        entityName = 'Defect'
    elif queryToken[0] == "U":
        entityName = 'HierarchicalRequirement'
    else:
        entityName = 'PortfolioItem'

    queryString = 'FormattedID = "' + queryToken + '"'

    if bPrintStatus: print ("Query = ", queryString)

    response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

    if response.resultCount == 0:
        errout('No item found for %s %s\n' % (entityName, arg))
        print ("Nothing")
    else:
        for item in response:
            print (item.FormattedID, "\t", item.Name)
            if hasattr(item, "TeamFeature"):
                print (returnAttrib(item, "TeamFeature.FormattedID"), "\t", returnAttrib(item, "TeamFeature.Name"))
                if hasattr(item.TeamFeature, "Parent"):
                    parent = item.TeamFeature.Parent
                    while (parent != None) and (parent.FormattedID[:3] != "PGM"):
                        #if (parent.FormattedID[0] == "F"):
                        #    outrow = ["FEA", parent.FormattedID, ] + \
                        #        [returnAttrib(parent, field, default="") for field in gFields]
                        #    print(outrow)
                        print (parent.FormattedID, "\t", parent.Name)
                        parent = parent.Parent

if __name__ == '__main__':
    main(sys.argv[1:])

