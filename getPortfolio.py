from pyral import Rally, rallySettings, rallyWorkset
import sys
import csv


"""
Get the whole tree for PGM362-> 13 PRJs -> FEA -> TF -> US
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

# mmmxxx
def FormRecordTeamFeature(tf):
    outrows = [["TF", tf.FormattedID, "" ] + [returnAttrib(tf, field, default="") for field in gFields]]
    if hasattr(tf, "UserStories"):
        for us in tf.UserStories:
            outrows += [["US", tf.FormattedID ] + FormRecordUserStory(us)]
    return outrows

def FormRecordUserStory(us):
    outrow = [us.FormattedID  ] + \
        [returnAttrib(us, field, default="") for field in gFields]
    return outrow

def ProcessPrograms(pgm, outfile):
    outrow = ["PGM", pgm.FormattedID, "","",  "", "" ] + \
        [returnAttrib(pgm, field, default="") for field in gFields]
    outfile.writerow(outrow)
    if hasattr(pgm, "Children"):
        ProcessProjects(pgm, outfile)

def ProcessProjects(pgm, outfile):
    for prj in pgm.Children:
        outrow = ["PRJ", pgm.FormattedID, prj.FormattedID, "", "", "" ] + \
            [returnAttrib(prj, field, default="") for field in gFields]
        outfile.writerow(outrow)
        if hasattr(prj, "Children"):
            ProcessFeatures(pgm.FormattedID, prj, outfile)


def ProcessFeatures(pgmFormattedID, prj, outfile):
    for fea in prj.Children:
        outrow = ["FEA", pgmFormattedID, prj.FormattedID, fea.FormattedID, "", "" ] + \
            [returnAttrib(fea, field, default="") for field in gFields]
        outfile.writerow(outrow)
        if hasattr(fea, "Children"):
            ProcessTeamFeatures(pgmFormattedID, prj.FormattedID, fea, outfile)


def ProcessTeamFeatures(pgmFormattedID, prjFormattedID, fea, outfile):
    if hasattr(fea, "Children"):
        for tf in fea.Children:
            outrow = ["TF", pgmFormattedID, prjFormattedID, fea.FormattedID, tf.FormattedID, "" ] + \
                [returnAttrib(tf, field, default="") for field in gFields]
            outfile.writerow(outrow)
            if hasattr(tf, "UserStories"):
                ProcessUserStories(pgmFormattedID, prjFormattedID, fea.FormattedID, tf, outfile)


def ProcessUserStories(pgmFormattedID, prjFormattedID, feaFormattedID, tf, outfile):
    for us in tf.UserStories:
        outrow = ["US", pgmFormattedID, prjFormattedID, feaFormattedID, tf.FormattedID,  us.FormattedID  ] + \
            [returnAttrib(us, field, default="") for field in gFields]
        outfile.writerow(outrow)


def returnAttrib(item, attr, default=""):
    locAttr = attr.split('.')
    if len(locAttr) == 1:
        return getattr(item, locAttr[0], default)
    else:
        return getattr(getattr(item, locAttr[0], ""), locAttr[1], default)

def printHelp():
    print ("    -h    Help")
            

def main(args):
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]
    #if len(args) != 1:
    #    errout('ERROR: Wrong number of arguments\n')
    #    printHelp()
    #    sys.exit(3)

    bPrintStatus = True
    if bPrintStatus: print ("options: ", options)
    if bPrintStatus: print ("args:    ", args)

    # Rally Settings
    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'  # tied to cme
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 


    if bPrintStatus: print ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

    fileName = 'GetP.csv'
    if bPrintStatus: print ("Printing to '%s'" % fileName)
    with open (fileName, 'w', newline='') as csvfile:
        #write header
        outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        #outrows = [["Type", "PGM", "PRJ", "FEA.ID", "TF.ID", "US.ID"] + [field for field in gFields]]
        outrows = [["Type", "TF.ID", "US.ID"] + [field for field in gFields]]
        #outfile.writerow(outrow)

        for arg in args:
            if arg[0] == "U":
                entityName = 'HierarchicalRequirement'
            else:
                entityName = 'PortfolioItem'

            #queryString = 'FormattedID = "PGM362"'
            queryString = 'FormattedID = "' + arg + '"'
            if bPrintStatus: print ("queryString: ", entityName, queryString)

            response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

            if response.resultCount == 0:
                errout('No item found for %s %s\n' % (entityName, arg))
            else:
                for resp in response:
                    if resp.FormattedID[:2] == "TF":
                        outrows += FormRecordTeamFeature(resp)
                    elif resp.FormattedID[:2] == "US":
                        outrows += [["US"] + FormRecordUserStory(resp)]
        outfile.writerows(outrows)


if __name__ == '__main__':
    main(sys.argv[1:])
