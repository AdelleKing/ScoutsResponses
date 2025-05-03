import pandas as pd
from jinja2 import Template
from docx import Document

# Load Google Form responses
df = pd.read_csv("responses.csv")

# Load the Word template
template_doc = Document("template.docx")

# Extract text from Word template
template_text = "\n".join([p.text for p in template_doc.paragraphs])

# Create a Jinja2 template
template = Template(template_text)

# Generate personalized documents
for index, row in df.iterrows():
    # Fill in placeholders using Jinja2
    rendered_text = template.render(
        TimeStamp=row["Timestamp"].upper(),
        UserEmail=row["Username"].upper(),
        ScoutGroup=row["Which group will your young person be joining?"].upper(),
        YPForename=row["Young Person's Forename(s)"].upper(),
        YPSurname=row["Young Person's Surname"].upper(),
        Nationality=row["Nationality"].upper(),
        School=row["School"].upper(),
        DateOfBirth=row["Date of Birth"].upper(),
        Gender=row["Gender"].upper(),
        YourTitle=row["Your Title"].upper(),
        PForename=row["Primary Parent/Guardian Forename(s)"].upper(),
        PSurname=row["Primary Parent/Guardian Surname"].upper(),
        YPRelationship=row["Relationship to young person"].upper(),
        Address=row["Address"].upper(),
        Postcode=row["Postcode"].upper(),
        TNumber=row["Telephone number"].upper(),
        SNumber=row["Second telephone number (work)"].upper(),
        PEmail=row["Primary Parent/Guardian Email Address"].upper(),
        Title=row["Title"].upper(),
        SForename=row["Secondary Parent/Guardian Forename(s)"].upper(),
        SSurname=row["Secondary Parent/Guardian Surname"].upper(),
        SRelationship=row["Secondary Parent/Guardian relationship to young person"].upper(),
        SAddress=row["Secondary Parent/Guardian Address"].upper(),
        SPostcode=row["Secondary Parent/Guardian Postcode"].upper(),
        STelephone=row["Secondary Parent/Guardian Telephone number"].upper(),
        AdditionalPhone=row["Additional telephone number (work)"],
        SEmail=row["Secondary Parent/Guardian Email Address"].upper(),
        Doctors=row["Doctors name and address"].upper(),
        Disabilities=row["Does the young person have any disabilities?"].upper(),
        DisabilitiesY=row["Disabilities: If Yes, please provide further information"],
        DietAllergies=row["Does the young person have any dietary needs or allergies?"].split(),
        DietAllergiesY=row["Dietary or Allergies:  If Yes, please provide further information"],
        Medical=row["Does the young person have any medical needs?"].split(),
        MedicalY=row["Medical needs:  If Yes, please provide further information"],
        Religion=row["Does the young person have a Religion or Faith?"].split(),
        ReligionY = row.get("Religion or Faith: If Yes, please provide further information", ""),
    )

    # Create a new Word document
    doc = Document()
    para = doc.add_paragraph()

    # Split text and apply bold formatting ONLY to dynamic values
    for word in rendered_text.split(" "):  
        run = para.add_run(word + " ")  # Add word to paragraph
        if word in [row["Username"], row["Which group will your young person be joining?"], row["Young Person's Forename(s)"], row["Young Person's Surname"]]:
            run.bold = True  # Apply bold formatting to dynamic data

    # Save the formatted document
    doc.save(f"output_{row['Young Person\'s Forename(s)']}_{row['Young Person\'s Surname']}.docx")

print("Documents generated successfully!")