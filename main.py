import pandas as pd
from docx import Document
from docx.shared import Pt
from flask import Flask, render_template, request, send_file, send_from_directory, url_for, current_app
from werkzeug.utils import redirect

app = Flask(__name__)

# @app.route("/")
# def homepage():
#     return render_template('temp.html', bg_class='classy')


@app.route("/", methods=['GET', 'POST'])
def generate_file():
    if request.method == 'POST':
        #Read the file
        new_file = request.files['uploaded_file']
        df = pd.read_csv(new_file)
        df.head()
        #Student Registered
        stu_reg = df.shape[0]
        # print(stu_reg)

        #Student did not select any session
        did_not_select_session = df[(df['Number of group sessions'] == 0) & (df['Number of 1:1 sessions'] == 0)].shape[0]
        # print(did_not_select_session)

        #Student select at least 1 session
        at_least_one_session_df = df[(df['Number of group sessions'] > 0) | (df['Number of 1:1 sessions'] > 0)]
        at_least_one_session = at_least_one_session_df.shape[0]
        # print(at_least_one_session)
        at_least_one_u2 = at_least_one_session_df[(at_least_one_session_df['School Year'] == 'Freshman') | (at_least_one_session_df['School Year'] == 'Sophomore')].shape[0]
        # print(at_least_one_u2)

        at_least_one_u3 = at_least_one_session_df[(at_least_one_session_df['School Year'] == 'Junior')].shape[0]
        # print(at_least_one_u3)

        # U4 include all the other school year that are not "Freshman, Sophomore and Junior"
        at_least_one_u4 = at_least_one_session_df[(at_least_one_session_df['School Year'] != 'Freshman') & (at_least_one_session_df['School Year'] != 'Sophomore') & (at_least_one_session_df['School Year'] != 'Junior')].shape[0]
        # print(at_least_one_u4)

        #Total Group Sessions Registered
        total_group_reg = at_least_one_session_df['Number of group sessions'].sum()
        # print(total_group_reg)
        unique_group_reg = at_least_one_session_df[at_least_one_session_df['Number of group sessions'] != 0].shape[0]
        # print(unique_group_reg)

        #Total group no show
        group_no_show = df[(df['Number of group sessions'] == df['Number of group no shows']) & (df['Number of group sessions'] != 0)]
        total_group_no_show = group_no_show['Number of group sessions'].sum()
        unique_group_no_show = group_no_show.shape[0]

        #Total 1:1 Sessions Registered
        one_one_total_reg = at_least_one_session_df['Number of 1:1 sessions'].sum()
        unique_one_one_reg = at_least_one_session_df[at_least_one_session_df['Number of 1:1 sessions'] != 0].shape[0]

        #Total 1:1 No Shows
        one_one_no_show = df[(df['Number of 1:1 sessions'] == df['Number of 1:1 no shows']) & (df['Number of 1:1 sessions'] != 0)]
        total_one_one_no_show = one_one_no_show['Number of 1:1 sessions'].sum()
        unique_one_one_no_show = one_one_no_show.shape[0]


        #Check in at least one session
        check_in_one_session_df = df[(df['Number of 1:1 sessions'] - df['Number of 1:1 no shows'] >0) | (df['Number of group sessions'] - df['Number of group no shows'] >0)]
        check_in_one_session = check_in_one_session_df.shape[0]
        check_in_one_u2 = check_in_one_session_df[(check_in_one_session_df['School Year'] == 'Freshman') | (check_in_one_session_df['School Year'] == 'Sophomore')].shape[0]
        check_in_one_u3 = check_in_one_session_df[(check_in_one_session_df['School Year'] == 'Junior')].shape[0]
        check_in_one_u4 = check_in_one_session_df[(check_in_one_session_df['School Year'] != 'Freshman') & (check_in_one_session_df['School Year'] != 'Sophomore') & (check_in_one_session_df['School Year'] != 'Junior')].shape[0]

        #No show to all sessions
        no_show_to_all_df = df[(df['Number of group sessions'] == df['Number of group no shows']) & (df['Number of 1:1 sessions'] == df['Number of 1:1 no shows']) & ((df['Number of group sessions'] != 0) | (df['Number of 1:1 sessions'] != 0))]
        no_show_to_all = no_show_to_all_df.shape[0]
        no_show_all_u2 = no_show_to_all_df[(no_show_to_all_df['School Year'] == 'Freshman') | (no_show_to_all_df['School Year'] == 'Sophomore')].shape[0]
        no_show_all_u3 = no_show_to_all_df[(no_show_to_all_df['School Year'] == 'Junior')].shape[0]
        no_show_all_u4 = no_show_to_all_df[(no_show_to_all_df['School Year'] != 'Freshman') & (no_show_to_all_df['School Year'] != 'Sophomore') & (no_show_to_all_df['School Year'] != 'Junior')].shape[0]

        #Writing the data to the document
        dc = Document()

        def paragraph_style(paragraph, bold=False):
            font = paragraph.runs[0].font  # Assuming each paragraph has only one run
            font.name = 'Calibri'
            font.size = Pt(10)
            font.bold = bold


        paragraphs = [
            f"Students Registered: {stu_reg}",
        f"                 Students tha did not select any sessions: {did_not_select_session}",
        f"                 Students tha select at least 1 session: {at_least_one_session}",
        f"                                 Class Level",
        f"                                                     U2:{at_least_one_u2}",
        f"                                                     U3:{at_least_one_u3}",
        f"                                                     U4:{at_least_one_u4}",
        f"                                 Total Group Sessions Registered: {total_group_reg}",
        f"                                                     By {unique_group_reg} unique students",
        f"                                 Total Group No Shows: {total_group_no_show}",
        f"                                                     By {unique_group_no_show} unique students (must have no showed to all group sessions",
        f"                                 Total 1:1 Sessions Registered: {one_one_total_reg}",
        f"                                                     By {unique_one_one_reg} unique students",
        f"                                 Total 1:1 No Shows: {total_one_one_no_show}",
        f"                                                     By {unique_one_one_no_show} unique students (must have no showed to all 1:1 sessions",
        f"                 Total unique students who checked-in for at leasr 1 session: {check_in_one_session}",
        f"                                 Class Level",
        f"                                                     U2:{check_in_one_u2}",
        f"                                                     U3:{check_in_one_u3}",
        f"                                                     U4:{check_in_one_u4}",
        f"                 Total unique students who no showed to all registered sessions: {no_show_to_all}",
        f"                                 Class Level",
        f"                                                     U2:{no_show_all_u2}",
        f"                                                     U2:{no_show_all_u3}",
        f"                                                     U2:{no_show_all_u4}"
        ]

        for pa in paragraphs:
            paragraph = dc.add_paragraph(pa)
            if "Students Registered" in pa or "Total unique students" in pa:
                paragraph_style(paragraph, bold=True)
            else:
                paragraph_style(paragraph)

        dc.save("Career Fair Report.docx")
        return send_file('Career Fair Report.docx', as_attachment=True)
    return render_template('temp.html', bg_class='classy')

if __name__ == "__main__":
    app.run(debug=True)
