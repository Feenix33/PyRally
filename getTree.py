from pyral import Rally, rallySettings, rallyWorkset
import sys
import csv

"""
Pull everything from the input point
everything = PGM > PRJ > FEA > TF > US

To Do:
    Defect as single entry
    Defects from one of the everything if available
    Only pull specific type from input flag
    Only pull down to specific type from input flag
    Change writing to file for speed
    Multiple input
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
    if len(args) != 1:
        errout('ERROR: Wrong number of arguments\n')
        printHelp()
        sys.exit(3)

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 

    bPrintStatus = True

    # form the query
    queryToken = args[0].upper()  #only search first one for now
    if queryToken[0] == "D":
        entityName = 'Defect'
    elif queryToken[0] == "U":
        entityName = 'HierarchicalRequirement'
    else:
        entityName = 'PortfolioItem'

    queryString = 'FormattedID = "' + queryToken + '"'
    fileName = queryToken + '.csv'
    fileName = 'out.csv'

    if bPrintStatus: print ("Query = ", queryString, fileName)

    if bPrintStatus: print ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

    response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

    if response.resultCount == 0:
        errout('No item found for %s %s\n' % (entityName, arg))
        print ("Nothing")
    else:
        if bPrintStatus: print ("Printing to '%s'" % fileName)
        with open (fileName, 'w', newline='', encoding="utf-8") as csvfile:
            #write header
            outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            outrow = ["Type", "PGM", "PRJ", "FEA.ID", "TF.ID", "US.ID"] + \
                        [field for field in gFields]
            outfile.writerow(outrow)

            for item in response:
                if queryToken[:2] == "PG":
                    processPGM(outfile, item)
                elif queryToken[:2] == "PR":
                    processPRJ(outfile, "", item)
                elif queryToken[:2] == "FE":
                    processFEA(outfile, "", "", item)
                elif queryToken[:2] == "TF":
                    processTF(outfile, "", "", "", item)
                elif queryToken[:2] == "US":
                    processUS(outfile, "", "", "", "", item)


def processPGM(outfile, pgm):
    outrow = ["PGM", pgm.FormattedID, "","",  "", "" ] + \
            [returnAttrib(pgm, field, default="") for field in gFields]
    outfile.writerow(outrow)
    if hasattr(pgm, "Children"):
        for prj in pgm.Children:
            processPRJ(outfile, pgm.FormattedID, prj)


def processPRJ(outfile, pgmID, prj):
    outrow = ["PRJ", pgmID, prj.FormattedID, "", "", "" ] + \
        [returnAttrib(prj, field, default="") for field in gFields]
    outfile.writerow(outrow)
    if hasattr(prj, "Children"):
        for fea in prj.Children:
            processFEA(outfile, pgmID, prj.FormattedID, fea)

def processFEA(outfile, pgmID, prjID, fea):
    outrow = ["FEA", pgmID, prjID, fea.FormattedID, "", "" ] + \
            [returnAttrib(fea, field, default="") for field in gFields]
    outfile.writerow(outrow)
    if hasattr(fea, "Children"):
        for tf in fea.Children:
            processTF(outfile, pgmID, prjID, fea.FormattedID, tf)

def processTF(outfile, pgmID, prjID, feaID, tf):
    outrow = ["TF", pgmID, prjID, feaID, tf.FormattedID, "" ] + \
        [returnAttrib(tf, field, default="") for field in gFields]
    outfile.writerow(outrow)
    if hasattr(tf, "UserStories"):
        for us in tf.UserStories:
            processUS(outfile, pgmID, prjID, feaID, tf.FormattedID, us)

def processUS(outfile, pgmID, prjID, feaID, tfID, us):
    outrow = ["US", pgmID, prjID, feaID, tfID,  us.FormattedID  ] + \
        [returnAttrib(us, field, default="") for field in gFields]
    outfile.writerow(outrow)


if __name__ == '__main__':
    main(sys.argv[1:])
