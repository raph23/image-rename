import csv
import datetime
import glob
import os
import platform
import pprint
import subprocess
import sys

# resources for opening .XLSX with Python --> eliminate the need to convert the file
# https://stackoverflow.com/questions/20105118/convert-xlsx-to-csv-correctly-using-python
# https://www.datacamp.com/community/tutorials/python-excel-tutorial

# TO DO: ADD output in csv file

# file extension for images and reports
IMAGE_FILE_EXTENSION = ".jpg"
REPORT_FILE_EXTENSION = ".csv"

RENAMED_FOLDER = "Final"
NOT_RENAMED_FOLDER = "Original"

def pause_program():
	pause_program = raw_input("\nPress ENTER to exit the program... ")

def open_image(file_name):
	if platform.system() == "Windows":
		os.startfile(file_name)
	else:
		subprocess.call(['open', file_name])
	
def sort_key(d):
    return d["file_name"]

def print_match(image, match):
	print("\nMatch found for image: %s" % (image["file_name"]))
	print("- Vendor SKU: %s" % (match["vendor_sku"]))
	print("- Vendor Product Name: %s" % (match["product"]))
	print("- Vendor Color: %s" % (match["vendor_color"]))
	print("- Style ID: %s\n" % (match["style_id"]))

def pick_reports():
	reports = []

	print("\n------------ PICK REPORT ------------\n")

	# look in the folder and add all images to the list
	for file_name in glob.glob("*" + REPORT_FILE_EXTENSION):
		reports.append(file_name)

	if len(reports) == 0:
		print('Did not find any reports in this folder!\nPlease make sure the report is in the same folder AND saved as .csv then try again!')
		pause_program()
		sys.exit()
	elif len(reports) == 1:
		report_name = 0
		print('Found 1 report in this folder, using "%s"' % reports[0])
	else:
		print("There are %s reports in this folder:" % len(reports))
		report_counter = 1
		for rep in reports:
			print("%s: %s" % (report_counter, rep))
			report_counter += 1
		print("")
		while True:
		    try:
		        picked_report = int(raw_input("Enter the number for the report you want to use: "))
		        report_name = picked_report-1
		    except ValueError:
		        print("Please enter a valid number (1-%i)." % len(reports))
		        continue
		    if picked_report < 1 or picked_report > len(reports):
		        print("Please enter a valid number (1-%i)." % len(reports))
		        continue
		    else:
		        #column pick was successfully parsed, exit the loop.
		        break

	with open(reports[report_name], 'rb') as f:
		reader = csv.reader(f)
		item_list = list(reader)

	del item_list[0:7]
	return item_list

def clean_name(cleaned_name):
	cleaned_name = cleaned_name.replace("_", " ")
	cleaned_name = cleaned_name.replace("-", " ")
	cleaned_name = cleaned_name.replace("/", " ")
	cleaned_name = cleaned_name.replace(".jpg", "")
	return cleaned_name

def parse_rows(report_rows):

	cleaned_rows = []

	for row in report_rows:
		
		row_dict = {}
		row_dict["style_id"] = row[4].lower()
		cleaned_name = clean_name(row[5].lower())
		row_dict["product"] = cleaned_name
		row_dict["vendor_sku"] = row[22].lower()
		cleaned_color = clean_name(row[23].lower())
		row_dict["vendor_color"] = cleaned_color
		row_dict["matched"] = "no"

		if row_dict not in cleaned_rows:
			cleaned_rows.append(row_dict)

	return cleaned_rows

def parse_images():
	cleaned_images = []

	for file_name in glob.glob("*" + IMAGE_FILE_EXTENSION):

		image_dict = {}
		image_dict["file_name"] = file_name

		cleaned_name = clean_name(file_name.lower())
		image_dict["cleaned_name"] = cleaned_name
		image_dict["angle"] = None
		image_dict["style_id"] = None
		image_dict["vendor_sku"] = None
		image_dict["product"] = None
		image_dict["matched"] = "no"
		image_dict["new_file_name"] = None

		cleaned_images.append(image_dict)

	sorted_images = sorted(cleaned_images, key=sort_key, reverse=False)

	return sorted_images

def adjust_angles(image_list, angle_list):

	print("\n------------ ANGLE ISSUES DETECTED ------------")
	set_with_angle_issues = []

	for angle in angle_list:
		if (angle["angle 1"] > 1 or angle["angle 2"] > 1 or angle["angle 3"] > 1) or (angle["angle 1"] > 0 and angle["angle 2"] == 0 and angle["angle 3"]> 0):
			set_with_angle_issues.append(angle)

	for issue_set in set_with_angle_issues:

		print('\nAngle mismatch for Style ID "%s":' % (issue_set["style_id"]))
		print("- Angle 1: %s images" % (issue_set["angle 1"]))
		print("- Angle 2: %s images" % (issue_set["angle 2"]))
		print("- Angle 3: %s images" % (issue_set["angle 3"]))

		for image in image_list:

			if issue_set["style_id"] == image["style_id"]:

				print("\nChange angle for %s?" % (image["file_name"]))
				print("Vendor SKU: %s" % (image["vendor_sku"]))
				print("Current angle: %s" % image["angle"])

				angle = pick_angle()
				if angle == 0:
					print("Skipped")
					image["new_file_name"] = None
					image["angle"] = None
					image["matched"] = "no"
				else:
					new_image_name = image["style_id"] + "_" + str(angle) + "_nocolor" + IMAGE_FILE_EXTENSION
					image["matched"] = "yes"
					image["new_file_name"] = new_image_name
					image["angle"] = angle

	check_angles(image_list)

def check_angles(image_list):

	angles_check_list = []
	mismatch = False

	for image in image_list:

		existed = False

		if image["style_id"] is not None:

			for image_angle in angles_check_list:

				if image["style_id"] in image_angle["style_id"]:

					if image["angle"] == 1:
						image_angle["angle 1"] += 1
					elif image["angle"] == 2:
						image_angle["angle 2"] += 1
					elif image["angle"] == 3:
						image_angle["angle 3"] += 1

					existed = True

			if existed == False:
				angle_dict = {}
				angle_dict["angle 1"] = 0
				angle_dict["angle 2"] = 0
				angle_dict["angle 3"] = 0
				angle_dict["style_id"] = image["style_id"]

				if image["angle"] == 1:
					angle_dict["angle 1"] += 1
				elif image["angle"] == 2:
					angle_dict["angle 2"] += 1
				elif image["angle"] == 3:
					angle_dict["angle 3"] += 1

				angles_check_list.append(angle_dict)

	for angle in angles_check_list:
		if (angle["angle 1"] > 1 or angle["angle 2"] > 1 or angle["angle 3"] > 1) or (angle["angle 1"] > 0 and angle["angle 2"] == 0 and angle["angle 3"]> 0):
			mismatch = True

	if mismatch == True:
		adjust_angles(image_list, angles_check_list)
	
	return image_list

def pick_angle():
	while True:
	    try:
	        angle = int(raw_input("What is the angle of the image? (1-3) or 0 to skip. "))

	    except ValueError:
	        print("Please enter a valid number.")
	        continue

	    if angle < 0 or angle > 3:
	        print("Please enter a valid number.")
	        continue
	    else:
	        #column pick was successfully parsed, exit the loop.
	        break
	return angle

def match_images_rows(image_list, row_list):

	print("\n------------ RENAME IMAGES ------------")

	for image in image_list:

		matches = []

		for row in row_list:

			if row["vendor_sku"] in image["cleaned_name"]:
				matches.append(row)

		if len(matches) > 1:
			for match in matches:
				if match["vendor_color"] in image["cleaned_name"]:
					
					print_match(image, match)

					open_image(image["file_name"])
					
					angle = pick_angle()
					if angle == 0:
						print("Skipped")
					else:
						new_image_name = match["style_id"] + "_" + str(angle) + "_nocolor" + IMAGE_FILE_EXTENSION
						image["matched"] = "yes"
						image["new_file_name"] = new_image_name
						image["angle"] = angle
						image["style_id"] = match["style_id"]
						image["vendor_sku"] = match["vendor_sku"]
						image["product"] = match["product"]

						for row in row_list:
							if row["style_id"] == match["style_id"]:
								row["matched"] = "yes"

		elif len(matches) == 1:

			print_match(image, matches[0])

			open_image(image["file_name"])

			angle = pick_angle()
			if angle == 0:
				print("Skipped")
			else:
				new_image_name = matches[0]["style_id"] + "_" + str(angle) + "_nocolor" + IMAGE_FILE_EXTENSION
				image["matched"] = "yes"
				image["new_file_name"] = new_image_name
				image["angle"] = angle
				image["style_id"] = matches[0]["style_id"]
				image["vendor_sku"] = matches[0]["vendor_sku"]
				image["product"] = matches[0]["product"]

				for row in row_list:
					if row["style_id"] == matches[0]["style_id"]:
						row["matched"] = "yes"
		# else:
		# 	print("\nNo matches found for image: %s" % (image["file_name"]))

	return image_list, row_list

def summarize_report(matched_rows):

	row_title_string = ("------------ STYLES IN REPORT WITHOUT IMAGES ------------\n")
	print(row_title_string)

	no_match = 0
	summary = []
	summary.append(row_title_string)

	for row in matched_rows:
		if row["matched"] == "no":
			row_string = ('Product: "%s" with Vendor SKU: %s and Style ID: %s' % (row["product"], row["vendor_sku"], row["style_id"]))
			print(row_string)
			summary.append(row_string)
			no_match += 1
	
	row_total_string = ("\nTotal rows NOT matched: %i" % (no_match))
	print(row_total_string)
	summary.append(row_total_string)

	return summary

def summarize_images(image_list):

	image_title_string = ("\n------------ IMAGE RENAME SUMMARY ------------\n")
	print(image_title_string)
	summary = []
	summary.append(image_title_string)
	match = 0
	no_match = 0

	for image in image_list:
		if image["matched"] == "yes":
			image_renamed_string = ("%s renamed to %s" % (image["file_name"], image["new_file_name"]))
			print(image_renamed_string)
			summary.append(image_renamed_string)
			match += 1
		else:
			image_not_renamed_string = ("%s not renamed" % (image["file_name"]))
			print(image_not_renamed_string)
			summary.append(image_not_renamed_string)
			no_match += 1

	image_total_matched_string = ("\nTotal images matched: %i" % (match))
	image_total_not_matched_string = ("Total images NOT matched: %i\n" % (no_match))
	print(image_total_matched_string)
	print(image_total_not_matched_string)
	summary.append(image_total_matched_string)
	summary.append(image_total_not_matched_string)

	return summary

def write_summary(images, report):

	x = datetime.datetime.now()
	file_name = "summary %s.txt" % (x.strftime("%m-%d-%y %I%M%p"))
	f = open(file_name,"w+")
	f.write("Renaming script run on: %s\n" % (x.strftime("%c")))

	for image_line in images:
		f.write("%s\n" % image_line)

	for row_line in report:
		f.write("%s\n" % row_line)

	f.close()

def rename_images(image_list):
	for image in image_list:
		if image["matched"] == "yes":
			new_file_name = ("%s/%s" % (RENAMED_FOLDER, image["new_file_name"]))
			os.rename(image["file_name"], new_file_name)
		else:
			new_file_name = ("%s/%s" % (NOT_RENAMED_FOLDER, image["file_name"]))
			os.rename(image["file_name"], new_file_name)

def make_folders():
	folders = [RENAMED_FOLDER, NOT_RENAMED_FOLDER]

	print("\n------------ CREATING FOLDERS ------------\n")

	for folder in folders:
		try:  
		    os.makedirs(folder)
		except OSError:
		    print('WARNING: Creation of the folder "%s" failed -- PLEASE CHECK IF THE FOLDER ALREADY EXISTS' % folder)
		else:  
		    print('Successfully created the folder "%s"' % folder)

make_folders()

report_rows = pick_reports()
cleaned_rows = parse_rows(report_rows)
cleaned_images = parse_images()

# pprint.pprint(cleaned_images)
# pprint.pprint(cleaned_rows)

matched_images, matched_rows = match_images_rows(cleaned_images, cleaned_rows)
checked_images = check_angles(matched_images)

# pprint.pprint(checked_images)
# pprint.pprint(cleaned_rows)

rename_images(checked_images)

image_summary = summarize_images(checked_images)
report_summary = summarize_report(matched_rows)
write_summary(image_summary, report_summary)

pause_program()