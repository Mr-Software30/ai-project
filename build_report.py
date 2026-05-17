from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

doc = Document()

# ── page margins ────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(2.5)

# ── helper functions ─────────────────────────────────────────
def add_para(text, bold=False, size=12, space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return p

def heading(text, level=1):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after  = Pt(6)
    return h

def page_break():
    doc.add_page_break()


# ════════════════════════════════════════════════════════════
#  PAGE 1 — COVER PAGE
# ════════════════════════════════════════════════════════════

# University name
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(60)
p.paragraph_format.space_after  = Pt(4)
run = p.add_run("University of Kurdistan Hewler")
run.bold = True
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0x1A, 0x3A, 0x6B)

# Department
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
run = p.add_run("Department of Computer Science and Engineering")
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0x1A, 0x3A, 0x6B)

# Module
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(60)
run = p.add_run("Artificial Intelligence Module")
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

# Main title
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(10)
run = p.add_run("Kurdistan Job Market Analysis")
run.bold = True
run.font.size = Pt(26)
run.font.color.rgb = RGBColor(0x1A, 0x3A, 0x6B)

# Subtitle
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(60)
run = p.add_run("Semester Project Report")
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

# Student info placeholder table
table = doc.add_table(rows=4, cols=2)
table.style = 'Table Grid'
labels = ["Student Name:", "Student ID:", "Instructor:", "Submission Date:"]
values = ["_______________________________",
          "_______________________________",
          "_______________________________",
          datetime.date.today().strftime("%B %Y")]

for i, (label, value) in enumerate(zip(labels, values)):
    row = table.rows[i]
    row.cells[0].text = label
    row.cells[1].text = value
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.size = Pt(11)

# Worth note
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(40)
run = p.add_run("Worth 25% of Module Grade")
run.bold = True
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)

page_break()


# ════════════════════════════════════════════════════════════
#  PAGE 2+ — REPORT CONTENT
# ════════════════════════════════════════════════════════════

# ── 1. Introduction ─────────────────────────────────────────
heading("1. Introduction", level=1)

add_para(
    "For this project I decide to collect data from a website called jobs.krd. "
    "This website is a job board that shows job vacancies in Kurdistan Region of Iraq. "
    "I think this is a good topic to study because jobs is something that affect everyone "
    "and there is not much research about the job market in Kurdistan. "
    "Also the website is public so anyone can see the data."
)
add_para(
    "The main goal of this project is to collect real job data, clean it, and then try some "
    "machine learning methods on it. I used Python for everything and put all the work "
    "in a Jupyter notebook like the project brief asked."
)


# ── 2. Data Collection ──────────────────────────────────────
heading("2. Data Collection", level=1)

add_para(
    "I used Python to scrape the data from jobs.krd. At first I tried to use Selenium "
    "because the website load the jobs using JavaScript and the normal requests library "
    "cannot see them. But then when I check the network traffic in the browser developer "
    "tools, I found that the website is calling a REST API to get the jobs data. "
    "So I use the requests library directly to call this API which is much faster and easier."
)
add_para(
    "The API address is: https://restapi.jobs.krd/api/v1/public/jobs"
)
add_para(
    "I collected 240 job listings from the website. The data was saved in a CSV file "
    "called jobs_krd.csv. Below is a description of the columns in the dataset:"
)

# Table of columns
col_table = doc.add_table(rows=9, cols=2)
col_table.style = 'Table Grid'
col_headers = ["Column", "Description"]
col_data = [
    ("title",     "The name of the job position"),
    ("company",   "The name of the company that posted the job"),
    ("location",  "The city where the job is located"),
    ("category",  "The type of work, for example Marketing or IT"),
    ("job_type",  "Full time, Part time, or Contract"),
    ("deadline",  "The last date to apply for the job"),
    ("salary",    "The salary if the company chose to show it"),
    ("url",       "The link to the full job page on jobs.krd"),
]

header_row = col_table.rows[0]
header_row.cells[0].text = col_headers[0]
header_row.cells[1].text = col_headers[1]
for cell in header_row.cells:
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = True
            run.font.size = Pt(11)

for i, (col, desc) in enumerate(col_data):
    row = col_table.rows[i + 1]
    row.cells[0].text = col
    row.cells[1].text = desc
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.size = Pt(10)

doc.add_paragraph()  # space after table


# ── 3. What I Did at Each Step ──────────────────────────────
heading("3. What I Did at Each Step", level=1)

# 3.1
heading("3.1  Data Preparation", level=2)
add_para(
    "First I load the CSV file into pandas and look at the data. I check for missing "
    "values and there was zero missing values which is good. I also check for duplicates "
    "and there was none. "
    "After that I did some cleaning:"
)
add_para("  •  I made a new column called salary_disclosed. If the company show the salary it become 1, if not it become 0.")
add_para("  •  I convert the deadline column from text to a real date using pandas.")
add_para("  •  I calculate how many days is left until the deadline and put it in a column called days_until_deadline.")
add_para(
    "Then I use LabelEncoder from sklearn to convert the text columns like job_type, location "
    "and category into numbers because machine learning models cannot work with text directly. "
    "At the end I split the data into 80% for training and 20% for testing."
)
add_para(
    "I also made two charts. The first one show that most jobs are Full time (226 out of 240). "
    "The second chart show that Erbil has the most jobs with 161 postings, followed by Duhok "
    "with 49 and Sulaymaniyah with 25."
)

# 3.2
heading("3.2  Regression", level=2)
add_para(
    "For the regression task I choose to predict the days_until_deadline column. This column "
    "is a real number so it is good for regression. I use Linear Regression from sklearn "
    "and the features I used are location, category, job type, and salary disclosed."
)
add_para(
    "The result was not very good. The MAE (Mean Absolute Error) was about 24 days which mean "
    "on average the model is wrong by 24 days. The R-squared was -0.024 which is basically zero "
    "or even negative. This mean the model is not able to predict the deadline from these features. "
    "I think the reason is that the deadline depends on each company decision, not on the type or "
    "location of the job. So there is no real pattern to learn."
)

# 3.3
heading("3.3  Classification", level=2)
add_para(
    "For classification I take the same days_until_deadline column and split it into two groups. "
    "If the deadline already passed (days is less than 0) I call it Expired and give it the label 0. "
    "If the deadline is today or in the future I call it Active and give it the label 1."
)
add_para(
    "I use a Decision Tree classifier with max depth 3. I also use class_weight balanced because "
    "the data is not balanced, 179 jobs are Expired and only 61 are Active. Without balancing the "
    "model just predict everything as Expired which is not useful even though it give 74% accuracy."
)
add_para(
    "The accuracy with the balanced Decision Tree was 62.5%. The model was able to correctly "
    "find 29 out of 34 Expired jobs but only 1 out of 14 Active jobs. The confusion matrix "
    "show this clearly."
)

# 3.4
heading("3.4  Clustering", level=2)
add_para(
    "For clustering I use K-Means algorithm. First I scale all the features using StandardScaler "
    "because K-Means use distance and if the columns have very different ranges it will not work correctly."
)
add_para(
    "I tried different values of k from 2 to 6 and plot the elbow curve. I choose k=4 because "
    "the inertia start to drop slower after that point. The four clusters I found were:"
)
add_para("  •  Cluster 0 (189 jobs): The main group. Full time jobs in Erbil, no salary shown.")
add_para("  •  Cluster 1 (12 jobs): Part time jobs, mostly in Erbil.")
add_para("  •  Cluster 2 (13 jobs): Jobs that show the salary. 100% of them disclose the salary.")
add_para("  •  Cluster 3 (26 jobs): Full time jobs in Sulaymaniyah.")
add_para(
    "The clustering found some interesting patterns even without us telling it what the groups "
    "should be. For example it separate the salary jobs and the Sulaymaniyah jobs automatically."
)

# 3.5
heading("3.5  Neural Network", level=2)
add_para(
    "For the neural network I build a simple MLP classifier using sklearn. The network have "
    "two hidden layers with 16 and 8 neurons. I use the adam optimizer and trained for up to "
    "1000 iterations but the model converge after 362 iterations."
)
add_para(
    "I use the same task as Section 3.3 so I can compare the results. The neural network got "
    "70.83% accuracy which is a bit better than the Decision Tree (62.5%). "
    "The loss curve show the model was learning because the loss go from 0.582 at the start "
    "down to 0.419 at the end."
)
add_para(
    "Important thing: I scale the features before training the neural network. "
    "This is very important because neural networks are very sensitive to the scale of the data."
)


# ── 4. Main Results ──────────────────────────────────────────
heading("4. Main Results", level=1)

add_para(
    "Here is a short summary of what I found from this project:"
)
add_para("  •  The job market in Kurdistan is mostly Full time jobs (94%) and mostly in Erbil (67%).")
add_para("  •  Very few companies show the salary, only 14 out of 240 jobs (about 6%).")
add_para("  •  The regression model did not work well because deadline dates don't depend on job type or location.")
add_para("  •  The classification task was hard because the data was imbalanced (75% Expired vs 25% Active).")
add_para("  •  K-Means clustering found 4 natural groups in the data: Erbil Full time, Part time, Salary-showing jobs, and Sulaymaniyah jobs.")
add_para("  •  The neural network did slightly better than the decision tree (70.83% vs 62.5%).")


# ── 5. What Did Not Work ──────────────────────────────────────
heading("5. What Did Not Work and What I Would Do Differently", level=1)

add_para(
    "The biggest problem in this project was the regression and classification tasks. "
    "The features I have (location, category, job type, salary) are not really connected "
    "to the deadline date. So both models gave low results. If I do this project again "
    "I would try to collect more useful features like the company size, the number of "
    "applicants, or the salary amount."
)
add_para(
    "The class imbalance was also a big problem. 75% of the jobs were already expired "
    "when I scraped the data. This make it very difficult for the classifier to learn "
    "the difference. Next time I would scrape the data at a better time, or I would "
    "use SMOTE or other techniques to balance the classes."
)
add_para(
    "Also I notice that the salary column was almost always 'Not disclosed' which make "
    "it not very useful as a feature. In the future I would try to get more data from "
    "different time periods to have more variety."
)
add_para(
    "Overall I think the project was a good learning experience. I learned how to "
    "collect real data from a website, clean it, and apply different machine learning "
    "methods. Even when the results are not perfect, it is useful to understand why "
    "the model is not working well."
)


# ── Save ─────────────────────────────────────────────────────
output_path = "/Users/m-strore/Documents/ai_project/report.docx"
doc.save(output_path)
print(f"Report saved to: {output_path}")
