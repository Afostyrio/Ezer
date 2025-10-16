from pptx import Presentation
from pptx.util import Inches
import os
import re

# Cleanup function DELETE LATER
os.remove("outputs/Presentación.pptx")

prs = Presentation()

with open("config", 'r') as instruction_file:
	instructions = instruction_file.readlines()
n_instructions = len(instructions)
location_match = re.compile("location: (.*)")
location = location_match.match(instructions[0]).group(1)
date_match = re.compile("date: (.*)")
date = date_match.match(instructions[1]).group(1)

def add_location_and_date(slide):
	txBox = slide.shapes.add_textbox(height=Inches(0.4), left=Inches(0), top=Inches(7), width=Inches(9))
	tf = txBox.text_frame
	tf.text = f"{location}, {date}"

def add_title_slide():
	slide = prs.slides.add_slide(prs.slide_layouts[0])
	slide.placeholders[0].text = "CEREMONIA DE PREMIACIÓN"
	add_location_and_date(slide)

def add_person_slide():
	slide = prs.slides.add_slide(prs.slide_layouts[3])
	add_location_and_date(slide)

def add_moment_slide():
	slide = prs.slides.add_slide(prs.slide_layouts[0])

# Main DSL parsing loop
for i in range(n_instructions):
	if re.match("title", instructions[i]):
		add_title_slide()
	if re.match("person", instructions[i]):
		add_person_slide()
	if re.match("moment", instructions[i]):
		add_moment_slide()


if not os.path.isdir("outputs"):
	os.mkdir("outputs")
prs.save("outputs/Presentación.pptx")

