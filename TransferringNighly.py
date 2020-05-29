import os
from os import path
import openpyxl
import pandas 
from shutil import copyfile
from settings import train_folder, test_TPFolder, global_passport_folder, global_passport_excel

class Transferring():
		
	# getting mds document name
	mds = input("Insert the MDS code!\n")
		
	def get_original_excel_file_path(self):
		# Getting the excel file with all the images with the path validation		
		while True:    
			original_excel_path = input("Insert the excel file path\n") 
			if '"' in original_excel_path:
				original_excel_path = original_excel_path.replace('"', '')
			if(path.exists(original_excel_path) is False):
				print("Wrong path, the excel file doesn't exist..")
			if(path.exists(original_excel_path) is True):  
				break  
		return original_excel_path
	
	def get_folder_path(self, folder_name):
		# Getting the drt folder where we have the images we want to split into train i test, with the path validation	
		while True:    
			folder_path = input("Insert your "+folder_name+" folder path\n")
			# if the path contains characters " they will be removed
			if '"' in folder_path:
				folder_path = folder_path.replace('"', '')	
			# if it's an invalid path raise an error
			if(path.exists(folder_path) is False):
				print("Wrong path, the folder doesn't exist..")
			if(path.exists(folder_path) is True):  
				break  
		return 	folder_path
				
	def split(self, excel_path, drt_path, train_path, test_TP_path):
	
		# open the original excel file we want to split
		original_workbook = pandas.read_excel(excel_path, index_col = False) 
		setup_frame = pandas.read_excel(excel_path, sheet_name= 'Test Setup', index_col = False) 
		# save total number of rows
		number_rows = len(original_workbook)	
		# mds folder inside the train folder
		mds_train_path = os.path.normpath(train_path+"/"+self.mds)
		# mds folder inside the test_TP folder
		mds_test_TP_path = os.path.normpath(test_TP_path+"/"+self.mds)
		# Excel file will be created in the train/mds folder 
		train_excel_path = os.path.normpath(mds_train_path+"/"+self.mds+".xlsx")
		# Excel file will be created in the test_TP/mds folder 
		test_TP_excel_path = os.path.normpath(mds_test_TP_path+"/"+self.mds+".xlsx")
		
		# checking if the mds path exists into the train folder, if not crete it
		if path.exists(mds_train_path) is False:
			try:
				os.mkdir(mds_train_path)
				print("mds Folder created in " + mds_train_path)
			except:
				print("Something wrong,folder couldn't be created in "+mds_train_path)
				quit()
				
		# checking if the mds path exists into the test_TP folder, if not crete it	
		if path.exists(mds_test_TP_path) is False:
			# checking if there are more than 100 images. If there are, create a test_TP folder otherwise don't.
			if (number_rows > 100):
				try:
					os.mkdir(mds_test_TP_path)
					print("mds Folder created in " + mds_test_TP_path)
				except:
					print("Something wrong, folder couldn't be created in "+mds_test_TP_path)
					quit()
		
		# Printing how many images we found in the excel file 
		print("Found " + str(number_rows) + " Images in the Excel file..\nCopying images...\n")
		# Iterating trough the excel rows getting the file name 
		for i in original_workbook.index:
			image_name = original_workbook['FileName'][i]			
			# if there are less than 100 images we just copy the imges from DRT folder to the train folder
			if(number_rows <= 100):		
				copyfile(drt_path+"/"+image_name , mds_train_path+"/"+image_name)
			# if there are between 100 and 200 images we copy the first 100 to the train folder  and those are left in the test_TP folder
			if(100 < number_rows < 200):		
				if(i < 100):
					copyfile(drt_path+"/"+image_name , mds_train_path+"/"+image_name)
				else:
					copyfile(drt_path+"/"+image_name  , mds_test_TP_path+"/"+image_name)
			# if there are more than 200 images we split half in the train folder and half in the test_TP one.
			elif(number_rows >= 200):
				if number_rows % 2 == 0:
					if(i < int(number_rows/2)):
						copyfile(drt_path+"/"+image_name , mds_train_path+"/"+image_name)
					else:
						copyfile(drt_path+"/"+image_name  , mds_test_TP_path+"/"+image_name)
				else:
					if(i < int(number_rows/2)):
						copyfile(drt_path+"/"+image_name , mds_train_path+"/"+image_name)
					else:
						copyfile(drt_path+"/"+image_name  , mds_test_TP_path+"/"+image_name)
		# delete frame
		del original_workbook
		print("Copy Complete!\n")
		print("--> Start new excels creation..")
		# Cheking how many images there are in the TruthData and in case there are less than 100, 
		# we copy the excel and the images in train folder 		
		if(number_rows <= 100):	
			train_workbook = pandas.read_excel(excel_path, index_col=False)	
			# change path column values into a folder where the images will be moved to
			train_workbook.loc[:, 'Path'] = mds_train_path+"\\"
			# re create HyperLink between the first and the second columns		
			train_workbook['Image Link'] = '=HYPERLINK("' + train_workbook['Path'] + train_workbook['FileName'] + '","' + train_workbook['FileName']+'")'
			# delete the first half amount of lines
			train_workbook = train_workbook.iloc[:number_rows]
			with pandas.ExcelWriter(train_excel_path, date_format='YYYY-MM-DD') as writer:
				train_workbook.to_excel(writer, sheet_name='Truth Data',  index=0)
				setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
			
			# delete the dataFrame
			del train_workbook
			
		# Cheking how many images there are in the TruthData and in case there are more than 100 and less than 200, 
		# we split the excel into half images for the first excel in train folder and half in the test folder 		
		elif(100 < number_rows <= 200):	
			train_workbook = pandas.read_excel(excel_path, index_col=False)	
			# change path column values into a folder where the images will be moved to
			train_workbook.loc[:, 'Path'] = mds_train_path+"\\"
			# re create HyperLink between the first and the second columns
			train_workbook['Image Link'] = '=HYPERLINK("' + train_workbook['Path'] + train_workbook['FileName'] + '","' + train_workbook['FileName']+'")'
			# delete the first half amount of lines
			train_workbook = train_workbook.iloc[:100]
			with pandas.ExcelWriter(train_excel_path, date_format='YYYY-MM-DD') as writer:
				train_workbook.to_excel(writer, sheet_name='Truth Data',  index=0)
				setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
			
			# delete the dataFrame
			del train_workbook
			
			test_workbook = pandas.read_excel(excel_path, index_col=False)
			# change path column values into a folder where the images will be moved to
			test_workbook.loc[:, 'Path'] = mds_test_TP_path+"\\"
			# re create HyperLink between the first and the second columns
			test_workbook['Image Link'] = '=HYPERLINK("' + test_workbook['Path'] + test_workbook['FileName'] + '","' + test_workbook['FileName'] + '")'
			# delete the half left amount of lines
			test_workbook = test_workbook.iloc[100-len(test_workbook):]
			with pandas.ExcelWriter(test_TP_excel_path, date_format='YYYY-MM-DD') as writer:
				test_workbook.to_excel(writer, sheet_name='Truth Data',  index=0)
				setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
			
			# delete the dataFrame
			del test_workbook
			
		# Cheking how many images there are in the TruthData and in case there are less than 200, 
		# we split the excel into 100 images for the first excel in train Folder 		
		elif(number_rows > 200):
			
			# Train Folder Excel
			train_workbook = pandas.read_excel(excel_path, index_col=False)	
			# change path column values into a folder where the images will be moved to
			train_workbook.loc[:, 'Path'] = mds_train_path+"\\"
			# re create HyperLink between the first and the second columns
			train_workbook['Image Link'] = '=HYPERLINK("' + train_workbook['Path'] + train_workbook['FileName'] + '","' +train_workbook['FileName']+'")'
			# Test_TP Folder Excel
			test_workbook = pandas.read_excel(excel_path, index_col=False)
			# change path column values into a folder where the images will be moved to
			test_workbook.loc[:, 'Path'] = mds_test_TP_path+"\\"
			# re create HyperLink between the first and the second columns
			test_workbook['Image Link'] = '=HYPERLINK("' + test_workbook['Path'] + test_workbook['FileName']+ '","' +test_workbook['FileName'] +'")'
			# delete the first half amount of lines
			train_workbook = train_workbook.iloc[:int(number_rows/2)]
			# if it's a even number of images, split half
			if(number_rows % 2 == 0):			
				# delete the half left amount of lines
				test_workbook = test_workbook.iloc[-int(number_rows/2):]
			else:
				test_workbook = test_workbook.iloc[-int(number_rows/2)-1:]

			with pandas.ExcelWriter(train_excel_path, date_format='YYYY-MM-DD') as writer:
				train_workbook.to_excel(writer, sheet_name='Truth Data',  index=0)
				setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
			# delete the dataFrame
			del train_workbook			
			
			with pandas.ExcelWriter(test_TP_excel_path, date_format='YYYY-MM-DD') as writer:
				test_workbook.to_excel(writer, sheet_name='Truth Data',  index=0)
				setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
			# delete the dataFrame
			del test_workbook
				
		print("Excels creation Complete!")
		return mds_train_path, mds_test_TP_path, number_rows, test_TP_excel_path, train_excel_path
		
	def copy_nightly_folder(self, mds_test_TP_path, test_TP_excel, eu_nightly_folder=None):		
		
		# open the original excel file we want to split
		workbook = pandas.read_excel(test_TP_excel, index_col=False)
		# get test setup frame
		setup_frame = pandas.read_excel(test_TP_excel, sheet_name='Test Setup', index_col=False)
		# create a folder if it doesn't exist
		if "PP" in self.mds:	
			#excel path
			if path.exists(global_passport_folder+"/"+self.mds) is False:
				os.mkdir(global_passport_folder+"/"+self.mds)
				print("mds Folder created in " + global_passport_folder)
			
			# copying all the images in the global Passport folder			
			for i in workbook.index:
				image_name = workbook['FileName'][i]
				copyfile(mds_test_TP_path+"/"+image_name , os.path.normpath(global_passport_folder+"/"+self.mds+"/"+image_name))
				
			#replace the images Path
			workbook.loc[:, 'Path'] = os.path.normpath(global_passport_folder+"/"+self.mds)+"\\"
			# Hyperlink Creation
			workbook['Image Link'] = '=HYPERLINK("' + workbook['Path'] + workbook['FileName'] + '","' + workbook['FileName']+'")'
			with pandas.ExcelWriter(global_passport_excel+'/'+self.mds+".xlsx", date_format='YYYY-MM-DD') as writer:
					workbook.to_excel(writer, sheet_name='Truth Data',  index=0)
					setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
			
		else:
			#excel path
			nighly_excel_path = os.path.normpath(eu_nightly_folder+"/"+self.mds+".xlsx")
					
			if path.exists(eu_nightly_folder+"/"+self.mds) is False:
				try:
					os.mkdir(eu_nightly_folder+"/"+self.mds)
					print("mds Folder created in " + eu_nightly_folder)
				except:
					print("Something wrong,folder couldn't be created in "+eu_nightly_folder)
					quit()
			# copying all the images in the eu nighly 			
			for i in workbook.index:
				image_name = workbook['FileName'][i]
				copyfile(mds_test_TP_path+"/"+image_name , os.path.normpath(eu_nightly_folder+"/"+self.mds+"/"+image_name))
				
			#replace the images Path
			workbook.loc[:, 'Path'] = os.path.normpath(eu_nightly_folder+"/"+self.mds)+"\\"
			# Hyperlink Creation
			workbook['Image Link'] = '=HYPERLINK("' + workbook['Path'] + workbook['FileName'] + '","' + workbook['FileName']+'")'
			with pandas.ExcelWriter(nighly_excel_path) as writer:		
				# save the excel file in train folder
				workbook.to_excel(writer, sheet_name='Truth Data',  index=0)	
				setup_frame.to_excel(writer, sheet_name='Test Setup',  index=0)
				
		# delete the dataFrame
		del workbook

		
# Main function					
if __name__ == "__main__":	
		
	# Class instance
	transfer = Transferring()

	# calling Class methods
	original_excel_path = transfer.get_original_excel_file_path()
	drt_folder = transfer.get_folder_path("DRT")
	nightly_folder = transfer.get_folder_path("nightly")
	mds_train_path, mds_test_TP_path, number_rows, test_TP_excel, train_excel_path = transfer.split(original_excel_path, drt_folder, train_folder, test_TPFolder)	
	if number_rows > 100: 
		print("--> Split " + drt_folder + " folder into " + mds_train_path + " and " + mds_test_TP_path)
	else:
		print("--> Moved images and from " + drt_folder + " to " + mds_train_path)
	print("--> Done. Transferring Complete!")
	print("--> Copying files in the nighly folder")
	if number_rows >= 200:
		if "PP" in transfer.mds:
			transfer.copy_nightly_folder(mds_test_TP_path, test_TP_excel)
		else:
			transfer.copy_nightly_folder(mds_test_TP_path, test_TP_excel, nightly_folder)
	else:
		if "PP" in transfer.mds:
			print("mds_train_path:"+mds_train_path, "\ntrain_excel_path:"+train_excel_path)
			transfer.copy_nightly_folder(mds_train_path, train_excel_path)
		else:
			transfer.copy_nightly_folder(mds_train_path, train_excel_path, nightly_folder)
	print("--> Nighly folder created correctly!")
	print("--> Entire transferring process successfully complete!")