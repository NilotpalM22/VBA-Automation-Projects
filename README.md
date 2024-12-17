# VBA-Automation-Projects
Collection of VBA scripts for workflow automation and data analysis.

## Cold Email Sender with Customization

This project contains a VBA macro that automates the process of sending cold emails to recruiters or companies. The user can customize the content of the email body, select the email's display or sending option, and attach a resume file to the email.

### Key Features:
- **Customizable Email Body**: The user can modify the content of the email before sending it, with placeholders for job title, company, and name.
- **Choose Display or Send**: Users can choose whether to display the email for review before sending or send it directly.
- **Resume Attachment**: Users are prompted to choose a resume file to attach to the email.
- **Job List Integration**: Emails are sent to multiple recruiters based on a list of job titles, companies, and email addresses from an Excel sheet.

### Files Included:
- **SendColdEmailsWithAttachment.xlsm**: The main Excel file containing the VBA macro to send emails.
- **UserForm**: A form that allows users to set their name, edit the email body, and choose the display/send options.

### How to Use:

1. **Open the Excel File**:
   - Open `SendColdEmailsWithAttachment.xlsm` in Excel.
   
2. **Enable Macros**:
   - Ensure macros are enabled in your Excel settings.

3. **Set Your Preferences**:
   - Open the UserForm by running the macro. The form will allow you to:
     - Enter your name (this will be used in the email signature).
     - Customize the body of the email.
     - Choose whether to display the email for review or send it directly.
     - Attach your resume by browsing for the file.

4. **Run the Macro**:
   - The email will be automatically generated for each recruiter/organization listed in the `Jobs` sheet.
   - If `Display` is selected, the email will open in Outlook for review before sending.
   - If `Send` is selected, the email will be sent directly to the email address.

### Code Overview:
- **VBA Script**: The script retrieves job information (Job Title, Company, Email) from a table in the `Jobs` sheet and personalizes the email content based on this data.
- **UserForm**: The form allows the user to edit the email body and select email settings (Display or Send).
- **Outlook Integration**: The code creates an email in Outlook and attaches the selected resume. It uses late binding to avoid version-specific issues with Outlook.

### Customization:
- You can modify the email body template to suit your needs.
- The code allows you to add or remove job-related fields from the `Jobs` sheet to tailor it to different types of outreach.

### Requirements:
- **Microsoft Outlook**: Required for sending emails through VBA.
- **Excel**: Ensure macros are enabled for the VBA code to run correctly.

### Contact:
For any questions or suggestions, feel free to reach out:
- **Email**: [mnilotpal@gmail.com](mailto:mnilotpal@gmail.com)
- **LinkedIn**: [LinkedIn Profile](https://www.linkedin.com/in/mnilotpal/)

---

Feel free to fork this repository and modify it to fit your needs. Contributions are always welcome!
