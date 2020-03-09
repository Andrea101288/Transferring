import os
from os import path
import openpyxl
import pandas as pandas
from shutil import copyfile

class Transferring():
		
	# getting mds document name
	mds = input("Insert the mds code! ex: MDS.2.0.PRT..ID.STD.012009.01\n")
		
	def getOriginalExcelFilePath(self):
		# Getting the excel file with all the images with the path validation		
		while True:    
			OriginalExcelPath = input("Insert the excel file path\n") 
			if '"' in OriginalExcelPath:
				OriginalExcelPath = OriginalExcelPath.replace('"', '')
			if(path.exists(OriginalExcelPath) is False):
				print("Wrong path, the excel file doesn't exist..")
			if(path.exists(OriginalExcelPath) is True):  
				break  
		return OriginalExcelPath
	
	def getDrtFolderPath(self):
		# Getting the drt folder where we have the images we want to split into train i test, with the path validation	
		while True:    
			drtPath = input("Insert your DRT folder path\n")
			# if the path contains characters " they will be removed
			if '"' in drtPath:
				drtPath = drtPath.replace('"', '')	
			# if it's an invalid path raise an error
			if(path.exists(drtPath) is False):
				print("Wrong path, the folder doesn't exist..")
			if(path.exists(drtPath) is True):  
				break  
		return 	drtPath	
		
	def getTrainFolderPath(self):
		# Getting the train folder where we want to move the first part of the images, with the path validation	
		while True:    
			trainPath = input("Insert your train folder path\n") 
			# if the path contains characters " they will be removed
			if '"' in trainPath:
				trainPath = trainPath.replace('"', '')	
				# if it's an invalid path raise an error
			if(path.exists(trainPath) is False):
				print("Wrong path, the folder doesn't exist..")
			if(path.exists(trainPath) is True):  
				break 
		return 	trainPath
	
	def gettest_TPFolderPath(self):
		# Getting the test folder where we want to move the second part of the images, with the path validation			
		while True:    
			test_TPPath = input("Insert your test_TP folder path\n") 
			# if the path contains characters " they will be removed
			if '"' in test_TPPath:
				test_TPPath = test_TPPath.replace('"', '')
				# if it's an invalid path raise an error
			if(path.exists(test_TPPath) is False):
				print("Wrong path, the folder doesn't exist..")
			if(path.exists(test_TPPath) is True):  
				break 
		return test_TPPath		
				
	def split(self, excelPath, drtPath, trainPath, test_TPPath):
	
		# getting mds document name
		# mds = input("Insert the mds code! ex: MDS.2.0.PRT..ID.STD.012009.01\n")
		# open the original excel file we want to split
		originalWorkBook = pandas.read_excel(excelPath, index_col = False) 
		# save total number of rows
		numberRows = len(originalWorkBook)	
		# mds folder inside the train folder
		mdsTrainPath = os.path.normpath(trainPath+"/"+self.mds)
		# mds folder inside the test_TP folder
		mdstest_TPPath = os.path.normpath(test_TPPath+"/"+self.mds)
		# Excel file will be created in the train/mds folder 
		trainExcelPath = os.path.normpath(mdsTrainPath+"/"+self.mds+".xlsx")
		# Excel file will be created in the test_TP/mds folder 
		test_TPExcelPath = os.path.normpath(mdstest_TPPath+"/"+self.mds+".xlsx")
		
		# checking if the mds path exists into the train folder, if not crete it
		if path.exists(mdsTrainPath) is False:
			try:
				os.mkdir(mdsTrainPath)
				print("mds Folder created in " + mdsTrainPath)
			except:
				print("Something wrong,folder couldn't be created in "+mdsTrainPath)
				quit()
		# checking if the mds path exists into the test_TP folder, if not crete it	
		if path.exists(mdstest_TPPath) is False:
			# checking if there are more than 100 images. If there are, create a test_TP folder otherwise don't.
			if (numberRows > 100):
				try:
					os.mkdir(mdstest_TPPath)
					print("mds Folder created in " + mdstest_TPPath)
				except:
					print("Something wrong, folder couldn't be created in "+mdstest_TPPath)
					quit()
		
		# Printing how many images we found in the excel file 
		print("Found " + str(numberRows) + " Images in the Excel file..")
		
		# Iterating trough the excel rows getting the file name 
		for i in originalWorkBook.index:
			imageName = originalWorkBook['FileName'][i]			
			# if there are less than 100 images we just copy the imges from DRT folder to the train folder
			if(numberRows < 100):		
				copyfile(drtPath+"/"+imageName , mdsTrainPath+"/"+imageName)
			# if there are between 100 and 200 images we copy the first 100 to the train folder  and those are left in the test_TP folder
			if(100 <= numberRows < 200):		
				if(i < 100):
					copyfile(drtPath+"/"+imageName , mdsTrainPath+"/"+imageName)
				else:
					copyfile(drtPath+"/"+imageName  , mdstest_TPPath+"/"+imageName)
			# if there are more than 200 images we split half in the train folder and half in the test_TP one.
			elif(numberRows > 200):
				if(i < int(numberRows/2)):
					copyfile(drtPath+"/"+imageName , mdsTrainPath+"/"+imageName)
				else:
					copyfile(drtPath+"/"+imageName  , mdstest_TPPath+"/"+imageName)
		# delete frame
		del originalWorkBook
		
		# Cheking how many images there are in the TruthData and in case there are less than 100, 
		# we copy the excel and the images in train folder 		
		if(numberRows <= 100):	
			trainWorkBook = pandas.read_excel(excelPath, index_col = False)	
			# change path column values into a folder where the images will be moved to
			trainWorkBook.loc[:, 'Path'] = mdsTrainPath+"\\"
			# re create HyperLink between the first and the second columns
			trainWorkBook['Image Link'] = '=HYPERLINK("' + trainWorkBook['Path'] + trainWorkBook['FileName'] +'")'
			# save the excel file in train folder
			trainWorkBook.to_excel(trainExcelPath, index=0)	
			# delete the dataFrame			
			del trainWorkBook
			
		# Cheking how many images there are in the TruthData and in case there are more than 100 and less than 200, 
		# we split the excel into half images for the first excel in train folder and half in the test folder 		
		if(100 < numberRows <= 200):	
			trainWorkBook = pandas.read_excel(excelPath, index_col = False)	
			# change path column values into a folder where the images will be moved to
			trainWorkBook.loc[:, 'Path'] = mdsTrainPath+"\\"
			# re create HyperLink between the first and the second columns
			trainWorkBook['Image Link'] = '=HYPERLINK("' + trainWorkBook['Path'] + trainWorkBook['FileName'] +'")'
			# delete the first half amount of lines
			trainWorkBook = trainWorkBook.iloc[:101]
			# save the excel file in train folder
			trainWorkBook.to_excel(trainExcelPath, index=0)
			# delete the dataFrame
			del trainWorkBook
			
			testWorkBook = pandas.read_excel(excelPath, index_col = False)
			# change path column values into a folder where the images will be moved to
			testWorkBook.loc[:, 'Path'] = mdstest_TPPath+"\\"
			# re create HyperLink between the first and the second columns
			testWorkBook['Image Link'] = '=HYPERLINK("' + testWorkBook['Path'] + testWorkBook['FileName'] +'")'
			# delete the half left amount of lines
			testWorkBook = testWorkBook.iloc[101-len(testWorkBook):]
			# save the excel file in test_TP folder
			testWorkBook.to_excel(test_TPExcelPath, index=0)
			# delete the dataFrame
			del testWorkBook
			
		# Cheking how many images there are in the TruthData and in case there are less than 200, 
		# we split the excel into 100 images for the first excel in train Folder 		
		if(numberRows > 200):	
			trainWorkBook = pandas.read_excel(excelPath, index_col = False)	
			# change path column values into a folder where the images will be moved to
			trainWorkBook.loc[:, 'Path'] = mdsTrainPath+"\\"
			# re create HyperLink between the first and the second columns
			trainWorkBook['Image Link'] = '=HYPERLINK("' + trainWorkBook['Path'] + trainWorkBook['FileName'] +'")'
			# delete the first half amount of lines
			trainWorkBook = trainWorkBook.iloc[:int(numberRows/2)]
			# save the excel file in train folder
			trainWorkBook.to_excel(trainExcelPath, index=0)
			# delete the dataFrame
			del trainWorkBook
			
			testWorkBook = pandas.read_excel(excelPath, index_col = False)
			# change path column values into a folder where the images will be moved to
			testWorkBook.loc[:, 'Path'] = mdstest_TPPath+"\\"
			# re create HyperLink between the first and the second columns
			testWorkBook['Image Link'] = '=HYPERLINK("' + testWorkBook['Path'] + testWorkBook['FileName'] +'")'
			# delete the half left amount of lines
			testWorkBook = testWorkBook.iloc[-int(numberRows/2)+1:]
			# save the excel file in test_TP folder
			testWorkBook.to_excel(test_TPExcelPath, index=0)
			# delete the dataFrame
			del testWorkBook
		
		return mdsTrainPath, mdstest_TPPath, numberRows
		
# Main function					
if __name__ == "__main__":
	
		
	# Class instance
	transfer = Transferring()
	# calling Class methods
	OriginalExcelPath = transfer.getOriginalExcelFilePath()
	drtFolder = transfer.getDrtFolderPath()
	trainFolder = transfer.getTrainFolderPath()
	test_TPFolder = transfer.gettest_TPFolderPath()
	mdsTrainPath, mdstest_TPPath, numberRows = transfer.split(OriginalExcelPath, drtFolder, trainFolder, test_TPFolder)	
	if numberRows > 100: 
		print("Split " + drtFolder + " folder into " + mdsTrainPath + " and " + mdstest_TPPath)
	else:
		print("Moved images and from " + drtFolder + " to " + mdsTrainPath)
	print("Done. Transferring Complete! ")