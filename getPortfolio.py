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

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 

    bPrintStatus = True

    if bPrintStatus: print ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

    if bPrintStatus: print ('Query execution...')

    #for arg in args:
    #    if arg[0] == "D":
    #        entityName = 'Defect'
    #    elif arg[0] == "U":
    #        entityName = 'HierarchicalRequirement'
    #    else:
    #        entityName = 'PortfolioItem'

    queryString = 'FormattedID = "PGM362"'
    entityName = 'PortfolioItem'

    if bPrintStatus: print ("Query = ", queryString)

    response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

    if response.resultCount == 0:
        errout('No item found for %s %s\n' % (entityName, arg))
    else:
        fileName = 'PGM362.csv'
        if bPrintStatus: print ("Printing to '%s'" % fileName)
        with open (fileName, 'w', newline='') as csvfile:
            #write header
            outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            outrow = ["Type", "PGM", "PRJ", "FEA.ID", "TF.ID", "US.ID"] + \
                        [field for field in gFields]
            outfile.writerow(outrow)

            for pgm in response:
                if pgm.FormattedID[:3] == "PGM":
                    outrow = ["PGM", pgm.FormattedID, "","",  "", "" ] + \
                        [returnAttrib(pgm, field, default="") for field in gFields]
                    outfile.writerow(outrow)

                    if hasattr(pgm, "Children"):
                        for prj in pgm.Children:
                            outrow = ["PRJ", pgm.FormattedID, prj.FormattedID, "", "", "" ] + \
                                [returnAttrib(prj, field, default="") for field in gFields]
                            outfile.writerow(outrow)

                            if hasattr(prj, "Children"):
                                for fea in prj.Children:
                                    outrow = ["FEA", pgm.FormattedID, prj.FormattedID, fea.FormattedID, "", "" ] + \
                                        [returnAttrib(fea, field, default="") for field in gFields]
                                    outfile.writerow(outrow)

                                    if hasattr(fea, "Children"):
                                        for tf in fea.Children:
                                            outrow = ["TF", pgm.FormattedID, prj.FormattedID, fea.FormattedID, tf.FormattedID, "" ] + \
                                                [returnAttrib(tf, field, default="") for field in gFields]
                                            outfile.writerow(outrow)

                                            if hasattr(tf, "UserStories"):
                                                for us in tf.UserStories:
                                                    outrow = ["US", pgm.FormattedID, prj.FormattedID, fea.FormattedID, tf.FormattedID,  us.FormattedID  ] + \
                                                        [returnAttrib(us, field, default="") for field in gFields]
                                                    outfile.writerow(outrow)


if __name__ == '__main__':
    main(sys.argv[1:])
