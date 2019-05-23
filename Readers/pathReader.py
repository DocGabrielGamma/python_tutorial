import sys, getopt

def getFilePaths(arguments):
   inputPath = ''
   outputPath = ''
   try:
      # h:Help , i: for input, o: for ouput
      opts, args = getopt.getopt(arguments,"hi:o:",["iPath=","oPath="]) 
   except getopt.GetoptError: 
      print ('Error incorrect arguments format it should be \n excelMerger.py -i <inputfile> -o <outputfile>')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print ('excelMerger.py -i <inputPath> -o <outputPath>')
         sys.exit()
      elif opt in ("-i", "--iPath"):
         inputPath = arg
      elif opt in ("-o", "--oPath"):
         outputPath = arg
   return (inputPath, outputPath)