from flask import Flask, request, send_file
import pandas as pd
import re

app = Flask(__name__)

def extract_name_from_filename(filename):
    match = re.match(r'Daily report_\d+_(.+)\.xls', filename)
    return match.group(1).replace('_', ' ') if match else None

def process_files(daily_reports, new_employee_file):
    interview_data = []
    new_employees = []

    for file in daily_reports:
        team_member = extract_name_from_filename(file.filename)
        df = pd.read_excel(file, engine="xlrd")
        for _, row in df.iterrows():
            interview_data.append({
                "Candidate Name": row["Candidate Name"].strip(),
                "Role": row["Role"].strip(),
                "Team Member": team_member
            })

    df_new = pd.read_excel(new_employee_file, engine="xlrd")
    new_employees.extend(df_new.to_dict(orient="records"))

    dashboard = []
    for emp in new_employees:
        for record in interview_data:
            if emp["Employee Name"].strip() == record["Candidate Name"] and emp["Role"].strip() == record["Role"]:
                dashboard.append({
                    "Employee Name": emp["Employee Name"],
                    "Join Date": emp["Join Date"],
                    "Role": emp["Role"],
                    "Team Member": record["Team Member"]
                })

    dashboard_df = pd.DataFrame(dashboard)
    output_file = "employee_dashboard.xlsx"
    dashboard_df.to_excel(output_file, index=False)
    
    return output_file

@app.route("/", methods=["GET", "POST"])
def upload_files():
    if request.method == "POST":
        daily_reports = request.files.getlist("daily_reports")
        new_employee_file = request.files["new_employee"]
        
        output_file = process_files(daily_reports, new_employee_file)
        
        return send_file(output_file, as_attachment=True)
    
    return '''
    <form method="post" enctype="multipart/form-data">
        Daily Report Files: <input type="file" name="daily_reports" multiple><br>
        New Employee File: <input type="file" name="new_employee"><br>
        <input type="submit" value="Upload & Process">
    </form>
    '''

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
