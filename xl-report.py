import xlwings as xw
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def generate_report():
    # Open the Excel workbook
    wb = xw.Book(r'C:/Users/kris/Desktop/VLM-VLS PM Calculator V5.5.xlsx')
    sheet = wb.sheets['Sheet1']

    # Get the data from the spreadsheet
    data = sheet.range('A1:C10').value

    # Create the PDF canvas
    pdf = canvas.Canvas('report.pdf', pagesize=letter)

    # Set the font and font size
    pdf.setFont('Helvetica', 12)

    # Write the data to the PDF
    y = 750
    for row in data:
        x = 50
        for col in row:
            pdf.drawString(x, y, str(col))
            x += 100
        y -= 20

    # Save the PDF
    pdf.save()

    # Close the workbook
    wb.close()

if __name__ == '__main__':
    generate_report()