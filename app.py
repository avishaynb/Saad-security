from flask import Flask, request, render_template
import pandas as pd
from datetime import datetime


app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    message = ''
    if request.method == 'POST':
        phone_number = request.form['phone_number']
        data = check_and_update_phone_number(phone_number)
        if data:
            f_name = data['PrivateName']
            l_name = data['LastName']
            last_checked = data['lastChecked']
            branch = data['Branch']
            return render_template('index.html', f_name=f_name, l_name=l_name, last_checked=last_checked, branch=branch,v_flag=True)
        else:
            message = "Phone number not found."
            return render_template('index.html', message=message)
        
    return render_template('index.html', message=message)
    

def check_and_update_phone_number(phone_number):
    # Path to your Excel file
    excel_file = 'מורשי כניסה - קיבוץ סעד.xlsx'
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)
    # Try to find the phone number in the DataFrame
    mask = df['phoneNumber'].apply(lambda x: str(x)) == phone_number[1:]
    found = df[mask]
    # Check if the phone number was found
    if not found.empty:
        # Update the lasttimecheckedDate for the found record
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        df.loc[mask, 'lastChecked'] = now
        # Write the updated DataFrame back to the Excel file
        # Ensure to use the 'openpyxl' engine for .xlsx files
        df.to_excel(excel_file, index=False, engine='openpyxl')
        # Return the found data as a string (or in another suitable format)
        return found.to_dict(orient='records')[0]
    else:
        return None



if __name__ == '__main__':

    app.run(debug=True)