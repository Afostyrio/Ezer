from pptx import Presentation
from pptx.util import Inches
import os
import re
import pandas as pd

# Cleanup function DELETE LATER
if os.path.isdir("outputs/Presentación.pptx"):
	os.remove("outputs/Presentación.pptx")

prs = Presentation()

with open("config", 'r') as instruction_file:
	instructions = instruction_file.readlines()
n_instructions = len(instructions)

medallistas_individual = pd.read_csv("inputs/csv/Medallistas Individual.csv", header=1)

# location_match = re.compile("location: (.*)")
# location = location_match.match(instructions[0]).group(1)
# date_match = re.compile("date: (.*)")
# date = date_match.match(instructions[1]).group(1)

# DSL functions
# def add_location_and_date(slide):
# 	txBox = slide.shapes.add_textbox(height=Inches(0.4), left=Inches(0), top=Inches(7), width=Inches(9))
# 	tf = txBox.text_frame
# 	tf.text = f"{location}, {date}"

def add_title_slide():
	slide = prs.slides.add_slide(prs.slide_layouts[0])
	slide.placeholders[0].text = "CEREMONIA DE PREMIACIÓN"

def add_person_slide(i: int):
	slide = prs.slides.add_slide(prs.slide_layouts[8])
	i += 1
	if i >= n_instructions: return None

	# Argument seeker
	while instructions[i][0] == "+":
		title_match = re.compile("\+ title: (.*)")
		title_get = title_match.match(instructions[i])
		if title_get: title = title_get.group(1)

		name_match = re.compile("\+ name: (.*)")
		name_get = name_match.match(instructions[i])
		if name_get: name = name_get.group(1)

		image_match = re.compile("\+ image: (.*)")
		image_get = image_match.match(instructions[i])
		if image_get: image = image_get.group(1)
		
		role_match = re.compile("\+ role: (.*)")
		role_get = role_match.match(instructions[i])
		if role_get: role = role_get.group(1)
		
		i+=1
		if i >= n_instructions: break
	
	if title: slide.placeholders[0].text = title
	if image: slide.placeholders[1].insert_picture(image)
	if name: slide.placeholders[2].text = name
	if role:
		add_role = slide.placeholders[2].text_frame.add_paragraph()
		add_role.text = role
		add_role.level = 1

def add_moment_slide(i: int):
	slide = prs.slides.add_slide(prs.slide_layouts[2])
	i += 1
	if i >= n_instructions: return None

	# Argument seeker
	while instructions[i][0] == "+":
		name_match = re.compile("\+ name: (.*)")
		name_get = name_match.match(instructions[i])
		if name_get: name = name_get.group(1)

		# image_match = re.compile("\+ image: (.*)")
		# image_get = image_match.match(instructions[i])
		# if image_get: image = image_get.group(1)
		
		i+=1
		if i >= n_instructions: break
	
	if name: slide.placeholders[0].text = name
	# if image: slide.placeholders[1].insert_picture(image)

# Main DSL parsing loop
for i in range(n_instructions):
	if instructions[i][0] == "#": continue # Ignore comments
	if re.match("title", instructions[i]):
		add_title_slide()
	if re.match("person", instructions[i]):
		add_person_slide(i)
	if re.match("moment", instructions[i]):
		add_moment_slide(i)

if not os.path.isdir("outputs"):
	os.mkdir("outputs")
prs.save("outputs/Presentación.pptx")
