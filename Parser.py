import pandas as pd
import os
from xlwt import Workbook
from datetime import date
from timeit import time

# Takes around 1min for 125 applications
start = time.time()
COLHEADER = 1
ROWHEADER = 0
BLANKCELL = ''
DEFAULTFASTTRACK = 'NO'

workbook = Workbook()
sheet1 = workbook.add_sheet('ParcerSheet1')

directory = "C:\\Users\\Friday\\Desktop\\Spring22\\Front Desk\\AppParser\\DENY'S TO BE UPDATED - March 2022, Part 1"
# directory = "C:\\Users\\Friday\\Desktop\\Spring22\\Front Desk\\AppParser\\TestFolder"
for row, eachApplication in enumerate(os.listdir(directory)):
    path = os.path.join(directory, eachApplication)
    if os.path.isfile(path):
        data = open(path, 'r')
        application = pd.read_html(data.name, index_col=False)
        # This list stores the information we retrieve from the application
        listOfEntries = []
        indexForTable = 0
        
        # APPLICATION INFORMATION
        appInfo = application[indexForTable]
        indexForTable += 1
        for indexForApplication in range(len(appInfo)):
            if appInfo.iloc[indexForApplication][ROWHEADER] == "Academic Plan":
                academicPlan = appInfo.iloc[indexForApplication][COLHEADER].split('__')[1]
            elif appInfo.iloc[indexForApplication][ROWHEADER] == "Email":
                emailAdd = appInfo.iloc[indexForApplication][COLHEADER]
            elif appInfo.iloc[indexForApplication][ROWHEADER] == "Application Submit Date":
                dateApplied = appInfo.iloc[indexForApplication][COLHEADER]
                
        # CONTACT INFORMATION
        contactInfo = application[indexForTable]
        indexForTable += 2
        for indexForContactInfo in range(len(contactInfo)):
            if contactInfo.iloc[indexForContactInfo][ROWHEADER] == "First Name":
                firstName = contactInfo.iloc[indexForContactInfo][COLHEADER]
            elif contactInfo.iloc[indexForContactInfo][ROWHEADER] == "Last Name":
                lastName = contactInfo.iloc[indexForContactInfo][COLHEADER]
            elif contactInfo.iloc[indexForContactInfo][ROWHEADER] == "Gender":
                gender = contactInfo.iloc[indexForContactInfo][COLHEADER]
                
        name = lastName+', '+firstName

        # ADDITIONAL INFORMATION
        additionalInfo = application[indexForTable]
        indexForTable += 1
        for indexForAdditionalInfo in range(len(additionalInfo)):
            if additionalInfo.iloc[indexForAdditionalInfo][ROWHEADER] == "Country of Citizenship":
                homeCountry = additionalInfo.iloc[indexForAdditionalInfo][COLHEADER]
            elif additionalInfo.iloc[indexForAdditionalInfo][ROWHEADER] == "Visa Permit Type":
                visaPermit = additionalInfo.iloc[indexForAdditionalInfo][COLHEADER]

        # EDUCATION INFORMATION
        for indexForEducation in range(indexForTable, len(application)):
            if application[indexForEducation].iloc[0][0] == "Education Information":
                # indexForEducation is the most recent education of the applicant
                indexForTable = indexForEducation
                break
        # TODO - There can be multple Education tables
        # Store them as master's degree if applicable

        educationInfo = application[indexForTable]
        indexForTable += 1
        for indexForEducationInfo in range(len(educationInfo)):
            if educationInfo.iloc[indexForEducationInfo][ROWHEADER] == "Institution Name":
                university = educationInfo.iloc[indexForEducationInfo][COLHEADER]
            elif educationInfo.iloc[indexForEducationInfo][ROWHEADER] == "Institution Name Other":
                college = educationInfo.iloc[indexForEducationInfo][COLHEADER]
            elif educationInfo.iloc[indexForEducationInfo][ROWHEADER] == "Degree":
                # TODO - If graduation data is later than today's date, append IP - In Progress to the degreeDesc
                if educationInfo.iloc[indexForEducationInfo][COLHEADER] == "Other" or pd.isna(educationInfo.iloc[indexForEducationInfo][COLHEADER]):
                    indexForEducationInfo += 1        
                degreeDesc = educationInfo.iloc[indexForEducationInfo][COLHEADER]
            elif educationInfo.iloc[indexForEducationInfo][ROWHEADER] == "Major":
                major = educationInfo.iloc[indexForEducationInfo][COLHEADER]
            elif educationInfo.iloc[indexForEducationInfo][ROWHEADER] == "Major GPA":
                gpa = round(float(educationInfo.iloc[indexForEducationInfo][COLHEADER]), 2)
            elif educationInfo.iloc[indexForEducationInfo][ROWHEADER] == "Education 2 Information":
                break

        # GRE SCORE
        for indexForGRE in range(indexForTable, len(application)):
            if application[indexForGRE].iloc[0][0] == "GRE Score":
                indexForTable = indexForGRE
                break
        greScore = application[indexForTable]
        indexForTable += 1
        for indexForGREScore in range(len(greScore)):
            if greScore.iloc[indexForGREScore][ROWHEADER] == "GRE General Verbal Score":
                verbalScore = greScore.iloc[indexForGREScore][COLHEADER]
            elif greScore.iloc[indexForGREScore][ROWHEADER] == "GRE General Quantitative Score":
                quantScore = greScore.iloc[indexForGREScore][COLHEADER]
            elif greScore.iloc[indexForGREScore][ROWHEADER] == "GRE General Writing Score":
                analyticalScore = greScore.iloc[indexForGREScore][COLHEADER]
        if pd.isna(verbalScore) or pd.isna(quantScore):
            totalGre = 0
        else:
            totalGre = int(verbalScore) + int(quantScore)
        # ENGLISH PROFICIENCY TEST
        EngTestScore = application[indexForTable]

        # TODO - verify which English proficiency test is taken
        engProficiencyTest = BLANKCELL
        for indexForEngScore in range(len(EngTestScore)):
            if EngTestScore.iloc[indexForEngScore][ROWHEADER] == "Duolingo English Test Score":
                engProficiencyTest = EngTestScore.iloc[indexForEngScore][COLHEADER]
            elif EngTestScore.iloc[indexForEngScore][ROWHEADER] == "TOEFL Total Score":
                engProficiencyTest = EngTestScore.iloc[indexForEngScore][COLHEADER]
            elif EngTestScore.iloc[indexForEngScore][ROWHEADER] == "IELTS Overall":
                engProficiencyTest = EngTestScore.iloc[indexForEngScore][COLHEADER]

        # Maintaining the order while framing a list of entries
        listOfEntries.append(dateApplied)
        listOfEntries.append(name)
        
        # BLANKCELL entries are to maintain the format of Excel sheet while copying it
        # to the original sheet
        dummyAttribute1 = BLANKCELL
        listOfEntries.append(dummyAttribute1)
        studentId = BLANKCELL
        listOfEntries.append(studentId)
        notes = BLANKCELL
        listOfEntries.append(notes)
        netId = BLANKCELL
        listOfEntries.append(netId)

        listOfEntries.append(emailAdd)

        firstEvaluator = BLANKCELL
        listOfEntries.append(firstEvaluator)
        levellingCourse = BLANKCELL
        listOfEntries.append(levellingCourse)
        reasonForDenial = BLANKCELL
        listOfEntries.append(reasonForDenial)

        listOfEntries.append(major)

        msMajor = BLANKCELL
        listOfEntries.append(msMajor)
        phDMajor = BLANKCELL
        listOfEntries.append(phDMajor)

        listOfEntries.append(university)
        listOfEntries.append(college)

        msUniversity = BLANKCELL
        listOfEntries.append(msUniversity)
        phDUniversity = BLANKCELL
        listOfEntries.append(phDUniversity)

        listOfEntries.append(degreeDesc)
        listOfEntries.append(verbalScore)
        listOfEntries.append(quantScore)
        listOfEntries.append(totalGre)
        listOfEntries.append(analyticalScore)

        bsPercentage = BLANKCELL
        listOfEntries.append(bsPercentage)
        bsPercentageQ = BLANKCELL
        listOfEntries.append(bsPercentageQ)

        listOfEntries.append(gpa)

        bsGPAQ = BLANKCELL
        listOfEntries.append(bsGPAQ)
        msPercentage = BLANKCELL
        listOfEntries.append(msPercentage)
        msGPA = BLANKCELL
        listOfEntries.append(msGPA)
        fastTrack = DEFAULTFASTTRACK
        listOfEntries.append(fastTrack)

        listOfEntries.append(engProficiencyTest)
        
        listOfEntries.append(academicPlan)
        listOfEntries.append(visaPermit)
        listOfEntries.append(homeCountry)
        listOfEntries.append(gender)
        #listOfEntries.append(date.today())

        # Run following code for each element in the list  
        for column, entry in enumerate(listOfEntries):
            sheet1.write(row+1,column+1, entry)
            workbook.save('ParserSheetMarchPart1.xls')
            # workbook.save('ParserSheetTest.xls')
end = time.time()
print("Runtime - ", end-start)     
'''
Pandora's box (alternative ways)
# workbook = pd.ExcelWriter('trailSheet.xls')
# emailAddress.to_excel(workbook, "sheet1", startcol = 3, startrow = 3)
# workbook.save()  
# Workbook is created
# from openpyxl import load_workbook
'''