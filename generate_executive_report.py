import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

EXCEL_FILE = "DocumentDataTracker.xlsx"
OUTPUT_REPORT = "Project_Status_Report.pdf"

def generate_pdf_report():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: {EXCEL_FILE} not found.")
        return
        
    df = pd.read_excel(EXCEL_FILE)
    
    # CHANGED: We are now pulling the clean, approved data that you just edited!
    approved_df = df[df["Review Status"] == "Approved"]
    
    if approved_df.empty:
        print("⚠️ No approved projects found to report. Make sure to change statuses to 'Approved' in Excel first!")
        return

    doc = SimpleDocTemplate(OUTPUT_REPORT, pagesize=letter,
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'ReportTitle', parent=styles['Heading1'], fontSize=22, leading=26,
        textColor=colors.HexColor("#1F4E79"), spaceAfter=6
    )
    
    meta_style = ParagraphStyle(
        'MetaText', parent=styles['Normal'], fontSize=10,
        textColor=colors.HexColor("#595959"), spaceAfter=20
    )

    # Header
    story.append(Paragraph("Master Project Progress Report", title_style))
    story.append(Paragraph("This executive brief aggregates all verified data extractions compiled by the automation pipeline.", meta_style))
    story.append(Spacer(1, 10))
    
    # Table Setup
    table_data = [["File Reference", "Project Name", "Objective Summary", "Status"]]
    
    for _, row in approved_df.iterrows():
        # Truncating objective strings so they wrap beautifully in the PDF grid
        objective = str(row["Objective"])
        short_obj = (objective[:45] + '...') if len(objective) > 45 else objective
        
        table_data.append([
            str(row["File Name"]),
            str(row["Project Name"]),
            short_obj,
            str(row["Status"])
        ])
        
    # Professional Corporate Layout Styling
    report_table = Table(table_data, colWidths=[130, 140, 160, 100])
    report_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E79")), # Navy Header
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E0E0")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FBFD")]), # Alternating subtle rows
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
    ]))
    
    story.append(report_table)
    doc.build(story)
    print(f"✔ Executive Report successfully built from your edits: {OUTPUT_REPORT}")

if __name__ == "__main__":
    generate_pdf_report()