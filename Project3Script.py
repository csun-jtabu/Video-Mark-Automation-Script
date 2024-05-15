# Jaztin Tabunda
# COMP 467 - Prof. Chaja
# 5-4-2024
# Project 3: "The Crucible"
#
# 1. Reuse Proj 1
# 2. Add argparse to input baselight file (--baselight),  xytech (--xytech) from proj 1
# 3. Populate new database with 2 collections: One for Baselight (Folder/Frames) and
# Xytech (Workorder/Location)
# 4. Download my amazing VP video,
# https://mycsun.box.com/s/v55rwqlu5ufuc8l510r8nni0dzq5qki7
# Links to an external site.
# 5. Run script with new argparse command --process <video file>
# 6. From (5) Call the populated database from (3), find all ranges only that fall
# in the length of video from (4)
# 7. Using ffmpeg or 3rd party tool of your choice, to extract timecode from video and
# write your own timecode method to convert marks to timecode
# 8. New argparse--output parameter for XLS with flag from (5) should export same CSV export as
# proj 1 (matching xytech/baselight locations), but in XLS with new column from files found
# from (6) and export their timecode ranges as well
# 9. Create Thumbnail (96x74) from each entry in (6), but middle most frame or closest to.
# Add to XLS file to it's corresponding range in new column
# 10. Render out each shot from (6) using (7) and upload them using API to frame.io
# (https://developer.frame.io/api/reference/)

import argparse  # used to perform file input/output
import pymongo  # used to upload to MongoDB
import shlex  # used when executing cli commands
import subprocess  # used to execute cli commands
import pandas  # used when converting to Excel file
import io  # need this to export to Excel file
import openpyxl  # need this to export to Excel file
from openpyxl.drawing.image import Image as OpenpyxlImage  # used to add images to excel file
from openpyxl.styles import Font  # this is used for formatting the text in excel file
from frameioclient import FrameioClient  # API used to upload to FrameIO website

# ----------------------------------------------------------------------------------------------------
# Project 1 Code:

rootDir = '/baselightfilesystem1/'  # global variable that tells the root directory we want to replace

# Assuming the Python script is in the same directory as the text file
# The file's text is saved into a string and returned
def importFileToString(fileName):
    file = open(fileName)
    fileText = file.read()
    return fileText
pass

# Method to turn each line in the text file into elements in a list
# This method is basically just a rename of splitlines()
def stringToList(string):
    lineList = string.splitlines()
    return lineList
pass

# ------------------ Methods to Help with pathConversion --------------------------------------

# Looking at a single line in Baselight_export.txt, we extract the path/directories not including the one we're trying
# to replace (rootdir). We return it as a string
# additionally, we'll ignore the numbers after the space which represent the frame numbers
def getEndDirectory(baseLightLine):
    global rootDir # this is the known directory we are trying to replace
    numToSkip = len(rootDir) # this is the num of characters of root directory
    newLine = baseLightLine[numToSkip:] # we cut out the root directory
    endDirectory = '' # we initialize the variable we'll be returning the new string in
    for char in newLine: # for each character in the line, we'll add the character to the endDirectory string
        if char.isspace(): # we'll only stop if we reach a space
            break          # this makes it so we don't add the frame numbers after
        else:
            endDirectory = endDirectory + char

    if endDirectory == '': # in case we don't add any end directories from the line
        return None
    else:
        return endDirectory
pass

# Looking at Xytech.txt, we'll check which location / new parent directory has the endDirectory we
# are passing in. The one that matches the endDirectory, we'll grab and send it back to the main program as a string
# Parameters: (List, String)
def getNewDirectory(lineListXytech, endDirectory):
    for line in lineListXytech: # checks each line in xytech
        if(line.find(endDirectory) != -1): # if there is a line that contains the end directory
            newDir = line.replace(endDirectory, '') # delete the ending directory
            return newDir # return the new directory as a string
pass

# ------------------------------------------------------------------------------------------

# This method will take in the entire text from baselight_export and xytech
# and will convert the path from baselight_export to the path stated in xytech
# The output will be a string with converted path
def pathConversion(bLightStr, xytechStr):
    global rootDir  # we call in the parent directory we are going to change
    bLightList = stringToList(bLightStr)  # we convert the file strings to a list of strings separated by new lines
    xytechList = stringToList(xytechStr)

    newBLightList = []  # this array will be used to store the newList after the conversion

    for line in bLightList:  # we check line in bLightList / each line in the string
        endDir = getEndDirectory(line)  # we search the line for the ending path we want
        if endDir != None:  # if there is an ending path on the line we looked at
            newDir = getNewDirectory(xytechList, endDir)  # we search the xytechList for the the line that contains the
                                                          # the path we want
            line = line.replace(rootDir, newDir) # self-explanatory
            newBLightList.append(line)

    delimiter = '\n' # this is what will be used to join the strings in the list back together
    newBLightStr = delimiter.join(newBLightList) # this converts the list back to a string

    return newBLightStr # we then return the string with the converted path
pass

# ------------------ Methods to Help with numConversion --------------------------------------

# This method gets the length of a given path from a line
# pass in a string/line and returns an integer
def pathLength(line):
    length = 0  # length starts at 0
    for char in line:  # it counts how many characters in the line
        length = length + 1
        if char == ' ':  # until we hit a space. (Since the lines we are dealing with have the paths at the start of
            break  # the line and there are no spaces in the file path, this should be good enough)
    return length

# This method looks at a line and returns a new line with only the numbers (i.e. new line with no path)
# This should work if the lines we are looking at have the ranges after the file path
def extractNums(line):
    dirLength = pathLength(line)  # this takes the pathLength
    newLine = line[dirLength:]  # we splice the line from the end of the path to the end of the string/line
    return newLine

# This method takes a string of numbers separated by spaces, converts them to actual values, and returns an integer list
def stringToNumList(line):
    stringList = list(line.split(" "))  # each number in the string is saved into a list
    numList = []  # this list will store the numbers in integer format
    for element in stringList:  # we look at each number in the stringList
        if((element != '<err>') and (element != '<null>') and ((element != ''))): # if it's not an error, a null, or
            numList.append(int(element)) # nothing, we add to the list as an integer value

    return numList

# This method will check each element in the passed-in numList to see if the consecutive numbers can be put in a range
# We're passing in a list with integer values and are returning a list with string values
def rangeChecker(numList):
    newNumList = []  # This is the resulting string list we'll be sending back
    startRange = numList[0]  # We initialize the first starting range with the first element in list
    currentRange = startRange  # The currentRange will be used to keep track of what element we're looking at in list
    endRange = startRange  # The endRange will be used to store the final value in a range
    for element in numList[1:]:  # we traverse through the list, skipping the first element
        if (element == (currentRange + 1)):  # if the current list element comes right after the currentRange
            endRange = element  # the endRange becomes the current element
            currentRange = endRange  # same goes for the currentRange
        else:  # if the current element doesn't come right after the currentRange
            if (startRange != endRange):  # If the range is more than a single element
                currentString = str(startRange) + '-' + str(endRange)  # the range will be saved as a string in stated
                newNumList.append(currentString)  # format and will be added to newNumList
            else:  # if the range is only a single element
                currentString = str(startRange)  # the range will be saved as a string in said format and will be added
                newNumList.append(currentString)  # to newNumList
            startRange = element  # now we reinitialize the loop by making the new starting range, current, and end
            currentRange = startRange  # the current element
            endRange = currentRange

    # Same as before, but this is for the last element/range in the list
    if (startRange != endRange):
        currentString = str(startRange) + '-' + str(endRange)
        newNumList.append(currentString)
    else:
         currentString = str(startRange)
         newNumList.append(currentString)

    return newNumList
pass

# This method gets all paths in a string and returns a list of strings
# Again, assuming all paths are at the beginning of their own line and have no spaces in them, this should work
# (This is for Project 1 Baselight file only)
def getAllPaths(string):
    stringList = stringToList(string)  # we first convert the string to a list
    newStringList = []  # we initialize a new list to hold the result
    for line in stringList:  # we check each line/element in the list
        dirLength = pathLength(line)  # we will get the length of the path in the current line
        newLine = line[0:dirLength-1]  # we will then splice the string from the start of the line to the path's length
        newStringList.append(newLine)  # we add the new result to the list
    return newStringList
pass

# ------------------------------------------------------------------------------------------

# This method will take in the string of baselight_export and will convert the frame numbers into ranges
# The output will be a list of strings
def numConversion(bLightStr):
    bLightList = stringToList(bLightStr) # we convert the string to a list of strings
    extractedFrameList = []  # this will store each list of frame numbers from each line from baselight_export
    rangeList = []  # this will store the newly converted frame ranges
    numList = []  # this will store integer values that were converted from strings
    pathList = []  # this will store the list of all paths from the bLightStr file
    newList = []  # this will store the final list with number conversions
    currentPathCount = 0  # this is a counter for the final for loop which will ensure an element in pathList with
                          # the same index will be combined with the element with the same index in rangeList

    for line in bLightList:  # for each line in the baselight_export
        newLine = extractNums(line)  # we remove the file paths and only keep the numbers in the line
        extractedFrameList.append(newLine)  # then we add the newly edited line to the extractedFrameList

    for line in extractedFrameList:  # we then check for each element/line in the extractedFrameList
        numList = stringToNumList(line)  # we convert the list of strings to a list of actual integer values
        newLine = rangeChecker(numList)  # we then check the newly made line/numList and edit the ranges
        rangeList.append(newLine)  # after all changes, we add it to a new list called rangeList

    pathList = getAllPaths(bLightStr)  # this will get all paths from bLightStr (remember: bLightStr is just baselight
                                       # in string format)

    for path in pathList:  # we're combining pathList and rangeList to come up with a new list
        currentList = rangeList[currentPathCount]  # we access the rangeList's current element which is a list of ranges
        for numRange in currentList:  # then for each numRange in the current element in the rangeList
            newLine = str(path) + ',' + str(numRange) # we make a new string/line containing new path and new range
            newList.append(newLine)  # we then append the result to newList
        currentPathCount = currentPathCount + 1  # after doing this for each numRange element, we move to the next index
                                                 # which
    return newList
pass

# ------------------ Methods to Help with assembleCSV --------------------------------------

# This method gets any other info we want from the xytech string/file. (producer/operator/job)
# Returns a string
def getXytechInfo(string, infoWanted):
    stringList = stringToList(string)  # We convert the xytech string to a list
    infoWanted = infoWanted + ': '  # This will be the string we will be looking for in the line
    length = len(infoWanted)  # We will get the length of the string of info we are looking for
    info = ''  # This will store the info we are looking for
    for line in stringList:  # for each line in the stringList
        if line.find(infoWanted) != -1:  # we'll check if the infoWanted is in that line.
            info = line[length:]  # if so, we'll store any information after that in the line
    return info
pass

# This method gets the notes from the Xytech workorder. Assuming the notes is always at the end of workorder
# Returns a string
def getNotes(string):
    stringList = stringToList(string) # We convert the xytech string to a list
    infoWanted = 'Notes:'  # We will look for this note
    notes = ''  # This will be where we store the notes
    flag = False  # This flag will be used to see if we passed the 'Notes:' keyword
    for line in stringList:  # for each line in the text/list we check if we passed Notes
        if flag == True:  # if we did already, then we append any text after it
            notes = notes + '\n' + line
        if line.find(infoWanted) != -1:  # if we find the 'Notes:' keyword we set the flag to True
            flag = True
    notes = notes[1:]  # we don't care about 'Notes:' so we cut it off
    return notes
pass

# ------------------------------------------------------------------------------------------

# This method is used to assemble all the new data into a CSV format according to spec
def assembleCSV(bLightList, xytechString):
    finalString = ''  # this is the new string we'll save the string in

    producer = getXytechInfo(xytechString, 'Producer')  # we get info for output from xytech string
    operator = getXytechInfo(xytechString, 'Operator')
    job = getXytechInfo(xytechString, 'Job')
    notes = getNotes(xytechString)

    line1 = 'Producer,Operator,Job,Notes\n'
    line2 = producer + ',' + operator + ',' + job + ',' + notes + '\n\n\n' # in csv format we assmble info we got

    line3 = 'Location,Frames to Fix,Timecode,Thumbnail\n'
    for line in bLightList:  # basically we are combining all the lines together to create one long string
        finalString = str(finalString) + str(line) + "\n"
    finalString = line1 + line2 + line3 + finalString  # this is the final formatted string

    return finalString
pass

# This writes the already CSV formatted string to a .csv file
def writeToCSV(string):
    fileName = 'project1Output.csv'  # .csv files are basically text files that have data values separated by commas
    file = open(fileName, 'w')  # we create and write to a new csv file
    file.write(string)
pass

# ------------------------------------------------------------------------------------------
# Argparse setup
parser = argparse.ArgumentParser(description='Used to manipulate data from files')

# These arguments are used to pass in the baselight and xytech files into the program (#2)
parser.add_argument('--baselight', dest='baselight', help='take in the file name for Baselight file')
parser.add_argument('--xytech', dest='xytech', help='take in the file name for Xytech file')

# This command/argument will be used to pass the name of the video and process the video
parser.add_argument('--process', dest='videoFile', help='take in the video file name to be processed')

# This command/argument will be used to output excel file of original Project 1 output with a catch
# The frame ranges within the video duration along with the new timecode ranges and thumbnails
parser.add_argument('--output', dest='outputFile', help='this will hold the output excel/xls file')

# This is where all the arguments from the parser will be stored
args = parser.parse_args()

# -----------------------------------------------------------------------------------------------
# Mongodb setup
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Project3DB"]  # database holding collections
baselightCol = mydb["BaselightCollection"]  # 1st collection
xytechCol = mydb["XytechCollection"]  # 2nd collection

# -----------------------------------------------------------------------------------------------
# sets up FrameIO client
frameIOclient = FrameioClient("fio-u-McuGX6hFbYCthgB1X55MM3c4fCU_mtCnPPC6-wWXHQm98iXbUqeNDUDrlq5DzsF_")  # my frame io access token
myFolder = "7c7e481d-94e2-4bd7-9e2d-e5179cee5efc"  # the location id i'm uploading the files to

# -----------------------------------------------------------------------------------------------

# This is used with the argparse command to get the baselight file we are working on
def getBaselightFile():
    global args
    if args.baselight != None:
        return args.baselight
    else:
        return None
pass

# This is used with the argparse command to get the xytech file we are working on
def getXytechFile():
    global args
    if args.xytech != None:
        return args.xytech
    else:
        return None
pass

# Input a list of baselight lines
# This is going to add to Baselight Collection (Folder/Frames)
def inputToBaselightCol(bLightList):
    global args, myclient, mydb, baselightCol

    for element in bLightList: # we check each file we pass in the command line
        line = element.split(",")
        myDict = {                 # each line's cells will be inputted
            'Folder': line[0],     # into a dictionary of Folder (path) and Frames (frame ranges)
            'Frames': line[1]
        }
        # This basically checks if the row we are entering is already in the Collection
        # i.e. duplicates
        dupeCheck = baselightCol.find_one(myDict)

        if dupeCheck:  # If there is a duplicate
            print('This is a duplicate. It\'s in BaselightCol already.')
        else:  # If it isn't we insert it into the Collection
            x = baselightCol.insert_one(myDict)
pass

# Input a list of Xytech lines
# This is going to add to Baselight Collection (Folder/Frames)
def inputToXytechCol(workorder, locationList):
    global args, myclient, mydb, xytechCol

    for element in locationList: # we check each file we pass in the command line
        myDict = {                 # each line's cells will be inputted
            'Workorder': workorder,     # into a dictionary Xytech Workorder and Location (new paths)
            'Location': element
        }
        # This basically checks if the row we are entering is already in the Collection
        # i.e. duplicates
        dupeCheck = xytechCol.find_one(myDict)

        if dupeCheck:  # If there is a duplicate
            print('This is a duplicate. It\'s in XytechCol already.')
        else:  # If it isn't we insert it into the Collection
            x = xytechCol.insert_one(myDict)
pass

# gets the Xytech workorder for db
# This returns a string containing the Xytech workorder
def getXytechWorkorder(string):
    stringList = stringToList(string)  # We convert the xytech string to a list
    infoWanted = 'Xytech Workorder '  # This will be the string we will be looking for in the line
    length = len(infoWanted)  # We will get the length of the string of info we are looking for
    info = ''  # This will store the info we are looking for
    for line in stringList:  # for each line in the stringList
        if line.find(infoWanted) != -1:  # we'll check if the infoWanted is in that line.
            info = line[length:]  # if so, we'll store any information after that in the line
    return info
pass

# gets the Xytech file paths for db
# This returns a list containing the Xytech file paths
def getXytechLocations(string):
    stringList = stringToList(string) # We convert the xytech string to a list
    infoWanted = 'Location:'  # We will look for this note
    pathList = []  # This will be where we store the notes
    flag = False  # This flag will be used to see if we passed the 'Notes:' keyword
    for line in stringList:  # for each line in the text/list we check if we passed Notes
        if flag == True and line == '':
            break
        if flag == True:  # if we did already, then we append any text after it
            pathList.append(line)
        if line.find(infoWanted) != -1:  # if we find the 'Notes:' keyword we set the flag to True
            flag = True
    return pathList
pass

# we get the number of frames the video has so when we get it from the database, we'll only get the ones
# with ranges and within the video max frame count
# Returns a string with the frame count of video
def getNumVideoFrames():
    string = ''  # this is going to be used to store the output of the command

    if args.videoFile != None:

        command = ("ffprobe "  # ffprobe used for extracting metadata
                   "-v error "  # this limits the output to just the frame number
                   "-select_streams v:0 "  # Only gets the first video stream it finds
                   "-show_entries stream=nb_frames "  # gets the number of frames metadata nb_frames
                   "-of default=nokey=1:noprint_wrappers=1 " +  # also limits the output to just the frame amount
                   args.videoFile)  # the video file we're referencing
        commandList = shlex.split(command)  # we split the arguments into elements in a list

        # Redirect STDERR to STDOUT, conjoining the two streams
        process = subprocess.Popen(commandList,  # args is the list of arguments we are passing
                                   stdout=subprocess.PIPE,  # stdout is the output we'll be producing
                                   stderr=subprocess.STDOUT,  # stderr is any errors that will be thrown from running this
                                   )

        for line in iter(process.stdout.readline, b''):  # b'' forces termination upon line with empty string
            string = string + line.decode().strip()  # we decode the bytes from process.stdout.readline and we take out any white space

    return string
pass

# calls populated database and gets xytech paths and baselight frames/paths
# this returns a list of strings/paths with converted xytech path and frame ranges
def getFromDatabase():
    global args, myclient, mydb, baselightCol, xytechCol

    videoFrameAmount = getNumVideoFrames()
    frameRangePattern = "^[0-9]*-[0-9]*$" # This makes sure that we are not including single frames (as stated in class)

    if args.videoFile != None:  # if --process is in commandline with a video file
        baseLightString = ''  # this is where we'll store baselight database info
        xytechString = ''  # this is where we'll store xytech database info

        # get paths and only ranges (i.e. no single frames)
        # additionally, it makes sure that the frame range doesn't exceed the video frame range
        for element in baselightCol.find({"Frames":{"$regex" : frameRangePattern}}): # regex checks for ranges
            rangeOnlyChecker = element["Frames"].split("-")  # we split the ranges into 2 values
            if(int(rangeOnlyChecker[1]) <= int(videoFrameAmount)):  # if the frame range is within the video range then we include it
                baseLightString = baseLightString + element["Folder"] + " " + element["Frames"] + "\n"

        for element in xytechCol.find():  # get the xytech file path
            xytechString = xytechString + element["Location"] + "\n"

        newString = pathConversion(baseLightString, xytechString)  # might as well convert the paths now since we
        pathList = stringToList(newString)  # called from both collections.
        return pathList  # we return the list with the new paths
pass

# This passes in a frame number and frames per second value in order to convert frame number to timecode
# It returns a string with a timecode
def frameToTC(frame, fps):
    widthTC = 2  # from the example in class the width of each HH, MM, SS, FF values are all 2
    FF = frame % fps  # the remainder of frame/fps will be the leftover frames
    SS = int(frame / fps)  # frame/fps will be how many seconds of frames we'll have and we'll round down
    MM = 0  # we initialize the minutes to zero
    HH = 0  # same for the hours

    if SS >= 60:  # instance where seconds are equal to or more than 60 seconds (convert to minutes)
        MM = int(SS / 60)  # for every 60 seconds we have a minute
        SS = int(SS % 60)  # remaining seconds from division will be the new seconds
    if MM >= 60: # instance where minutes are equal to or more than 60 seconds (convert to hours)
        HH = int(MM / 60)  # for every 60 minutes we have an hour
        MM = int(MM % 60)  # remaining minutes from division will be the new minutes

    FF = f"{FF:0{widthTC}d}"  # format each so it's always width = 2. (0-9)(0-9)
    SS = f"{SS:0{widthTC}d}"
    MM = f"{MM:0{widthTC}d}"
    HH = f"{HH:0{widthTC}d}"

    TC = HH + ":" + MM + ":" + SS + ":" + FF  # assemble the newly converted timecode as string

    return TC  # return the timecode
pass

# This method will convert and append the timeCode to the end of the current line
# which would contain file path and ranges. additionally, it preps for output file
def addTC(pathList):
    newPathList = []
    fps = 60
    if args.videoFile != None:
        for line in pathList:
            lineList = line.split(' ')  # separate the file path from ranges
            rangeList = lineList[1].split('-')  # this will contain the frame ranges. we split so we can get both values
            initialTC = frameToTC(int(rangeList[0]), fps)
            finalTC = frameToTC(int(rangeList[1]), fps)
            newLine = line + " " + str(initialTC) + "-" + str(finalTC)
            newLine = newLine.replace(" ", ",")
            newPathList.append(newLine)
    return newPathList
pass

# This generates the Excel file from the comma delimited string
def generateExcelFile(csvString):
    if args.outputFile != None:
        print("Generating Excel File")
        frame = pandas.read_csv(io.StringIO(csvString))  # we make a csv object that is read by pandas
        frame.to_excel(args.outputFile, index=False)  # we export to excel file
pass

# This gets the middle most frame from a range by taking in the pathList which contains the ranges
# This will return a list with the middlemost value in each range.
def getMiddleMostFrame(pathList):
    middle = 0  # current middle initialized
    middleList = []  # for each range, we'll store the middle value in this list
    if args.outputFile != None:
        for line in pathList:
            lineList = line.split(' ')  # separate the file path from ranges
            rangeList = lineList[1].split('-')  # this will contain the frame ranges. we split so we can get both values
            middle = ((int(rangeList[1]) + int(rangeList[0])) / 2)  # median formula to find the middle value
            middleList.append(middle)  # append the middle value in the range to the list
    return middleList
pass

# This will generate the thumbnails we need to put in the excel file
# Additionally, it will return the range list so we can segway into
def generateThumbnails(pathList):
    if args.videoFile != None:
        middleList = getMiddleMostFrame(pathList)  # we get a list of the middlemost frames in each range
        rangeList = []  # we'll be storing all the ranges here, so we can easily reference the thumbnails later
        counter = 0  # this is the counter to access each row
        print("Generating Thumbnails:")
        for middleValue in middleList:
            print("Middle Thumbnail: " + str(int(middleValue)))
            range = pathList[counter].split(" ")  # this will separate the path from the frame range
            rangeList.append(range[1])  # this will append the range to a new list we'll be returning
            middleValue = int(middleValue)  # This is to ensure we are working with an integer value
            middleValue = middleValue - 1  # The command we are using starts at an index 0 so this is to offset

            # this is how we specify the frame image we want and we resize it per spec
            vfStatement = "\"select=eq(n\," + str(middleValue) + "\"),scale=96:74"
            outputFileName = "ThumbnailRange" + range[1] + ".png"  # the name of the output file

            command = ("ffmpeg "  # we use ffmpeg 
                       "-i " + str(args.videoFile) +  # this is the file we are inserting
                       " -vf " + str(vfStatement) +  # this is to filter. we pass our query that gets the frame image we want
                       " -vframes 1" +  # this ensures 1 image to 1 output file
                       " -y " + str(outputFileName))  # the name of the output file

            commandList = shlex.split(command)  # we split the arguments into elements in a list

            # Redirect STDERR to STDOUT, conjoining the two streams
            process = subprocess.Popen(commandList,  # args is the list of arguments we are passing
                                       stdout=subprocess.PIPE,  # stdout is the output we'll be producing
                                       stderr=subprocess.STDOUT,
                                       # stderr is any errors that will be thrown from running this
                                       )
            process.wait()  # so we wait until the the last thumbnail is generated before moving on to the next
            counter = counter + 1  # move to the next line

        return rangeList  # return the list of ranges to easily reference thumbnails
pass

# This method is used to add the images to the Excel file and to format the columns/rows
def addThumbnailsToExcel(rangeList):
    if args.outputFile != None:
        print("Adding Thumbnails to Excel")
        excelFile = openpyxl.load_workbook("Project3Output.xlsx")  # the excel file we are inserting the images onto
        excelSheet = excelFile.active  # we set the sheet we are currently looking at from the file
        excelSheet['A3'].font = Font(bold=True) # just to format the cells to be bold
        excelSheet['B3'].font = Font(bold=True)
        excelSheet['C3'].font = Font(bold=True)
        excelSheet['D3'].font = Font(bold=True)
        excelSheet.insert_rows(3, amount=1)  # new row/space
        counter = 5  # the thumbnail row starts at row 4
        for givenRange in rangeList:
            cellReference = 'D' + str(counter)  # this will be the cell we are referencing
            currentCell = excelSheet[cellReference]  # we set the cell we are referencing on the sheet
            if currentCell.value is None:  # we check if there's something in the cell, if not, we're good to add the thumbnails
                imageName = 'ThumbnailRange' + givenRange + '.png'  # name of the image file
                formattedImage = OpenpyxlImage(imageName)  # how we open the image
                excelSheet.add_image(formattedImage, cellReference)  # we add the image to the excel sheet
                excelSheet.row_dimensions[counter].height = 60
                counter = counter + 1  # this is going to reference the next row
        excelSheet.column_dimensions['A'].width = 50 # this is just formatting
        excelSheet.column_dimensions['B'].width = 13
        excelSheet.column_dimensions['C'].width = 22
        excelSheet.column_dimensions['D'].width = 40
        excelFile.save("Project3Output.xlsx")  # save the rewritten excel file
pass

# We need this so ffmpeg can understand what time frame we're talking about
# ffmpeg needs the timecode to be in this format: HH:MM:SS.MS
# String of new timecode is returned
def timeCodeToTCMS(timeCode, fps):
    widthTC = 2  # from the example in class the width of each HH, MM, SS, FF values are all 2
    currentTimecode = timeCode.split(":")  # we split up the timecode using the colon as a delimiter

    HH = currentTimecode[0] # save everything from previous timecode
    MM = currentTimecode[1]
    SS = currentTimecode[2]
    FF = currentTimecode[3]

    MS = (int(FF) * 1000) / fps  # calculates milliseconds
    MS = int(MS)  # rounds down to nearest integer

    formattedMS = str(MS).zfill(3)

    TC = HH + ":" + MM + ":" + SS + "." + formattedMS  # assemble the newly converted timecode as string

    return TC  # return the timecode
pass

# This generates the renders/videos
def generateRenders(listWTC):

    if args.videoFile != None:
        print("Generating Renders")
        renderNameList = []
        for timeCodeRange in listWTC:
            range = timeCodeRange.split(",")  # this will separate the path from the frame range
            currentTCRange = range[2].split("-")  # this will separate the range into 2 values
            start = timeCodeToTCMS(currentTCRange[0],60)  # convert first part of range timecode
            end = timeCodeToTCMS(currentTCRange[1], 60)  # convert second part of range timecode
            outputFileName = "Clip" + range[1] + ".mp4"  # the name of the output file

            # ffmpeg -i <InputName> -ss <InsertStartTimeCode> -to <InsertEndTimeCode> -y -c:a copy <OutputName>
            command = ("ffmpeg "  # we use ffmpeg 
                       "-i " + str(args.videoFile) +  # this is the file we are inserting
                       " -ss " + start +  # this is to filter. we pass our query that gets the frame image we want
                       " -to " + end + # this ensures 1 image to 1 output file
                       " -y -c:a copy " + str(outputFileName))  # the name of the output file

            commandList = shlex.split(command)  # we split the arguments into elements in a list

            # Redirect STDERR to STDOUT, conjoining the two streams
            process = subprocess.Popen(commandList,  # args is the list of arguments we are passing
                                       stdout=subprocess.PIPE,  # stdout is the output we'll be producing
                                       stderr=subprocess.STDOUT, # stderr is any errors that will be thrown from running this
                                       )

        return renderNameList  # return the list of ranges to easily reference thumbnails
pass

# Pretty self explanatory. Used to upload to FrameIO using frameioclient api
def uploadToFrameIO(rangeList):
    global frameIOclient, myFolder
    if args.outputFile != None:
        print("Uploading to FrameIO")
        for givenRange in rangeList:
            videoName = 'Clip' + givenRange + '.mp4'
            frameIOclient.assets.upload(myFolder, videoName)  # (FrameioURLPath, filePathToUpload/fileNameToUpload)
pass

def main():

    # this is how we'll start the conversion process. we'll basically pass in the file names
    # using argparse
    baselightFileName = getBaselightFile()
    xytechFileName = getXytechFile()
    bLightString = ''
    xytechString = ''

    # This is the pre processing
    # Reminder from Project1, we did most of the conversions as a string
    try:
        bLightString = importFileToString(baselightFileName)
    except:
        print("No proper Baselight file passed")
    try:
        xytechString = importFileToString(xytechFileName)
    except:
        print("No proper Xytech file passed")

    # This is going to make sure that the end of the string/file has no whitespace.
    bLightString = bLightString.rstrip()  # (white space breaks the range method)
    blightList = numConversion(bLightString)  # we do the frame conversion only since that's what we need for collection
    inputToBaselightCol(blightList)  # we input the baselight info into mongodb (Folder/Frames)

    workorder = getXytechWorkorder(xytechString)  # we get the xytech workorder for database upload
    locationList = getXytechLocations(xytechString)  # we get the file paths for database upload
    inputToXytechCol(workorder, locationList)  # we input the xytech info into mongodb (Workorder/Location)

    listOfPaths = getFromDatabase()  # database call to get and convert xytech and baselight paths from db
    listOfPathsTC = addTC(listOfPaths)  # we add the timecodes to the list of paths we already have
    csvFormattedString = assembleCSV(listOfPathsTC, xytechString)  # this assembles the comma delimited string

    generateExcelFile(csvFormattedString)  # generates the excel file
    rangeList = generateThumbnails(listOfPaths)  # generates the thumbnails from video
    addThumbnailsToExcel(rangeList)  # adds the thumbnails to excel file
    generateRenders(listOfPathsTC)  # creates the clips/videos
    uploadToFrameIO(rangeList)  # uploads to frameIO
pass

main()