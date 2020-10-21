"""
Script to automate the data entry and pdf conversion of cover letters

@author Harsh Padhye
"""
import shutil, os
from mailmerge import MailMerge
from datetime import date
from docx2pdf import convert

template = "Harsh Padhye Cover Letter Template.docx"

document = MailMerge(template)

# ask user for Company/Job details
company_name = input("Name of Company: ")
job_title = input("Job Title: ")
job_skill = input("Main Job Skill: ")

# replace all MailMerge values
document.merge(
    company=company_name,
    job=job_title,
    skill=job_skill,
    date='{:%B %d, %Y}'.format(date.today())
)
outfile = f"{company_name} Cover Letter"

# write the merged version to a new document
document.write(outfile + ".docx")

#convert to pdf
convert(outfile + ".docx")

# move to appropriate folder in "Job Resources"
shutil.move(outfile + ".pdf","C:\\Users\\harsh\\Desktop\\Job Resources\\Cover Letters")
shutil.move(outfile + ".docx", "C:\\Users\\harsh\\Desktop\\Job Resources\\Cover Letters\\docs")
