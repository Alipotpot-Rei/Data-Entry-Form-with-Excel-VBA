# Automated Data Entry Form with Email Submission (Excel VBA)

### Project Overview

The project is an Excel-based data entry form with a VBA-powered "Submit" button that automatically sends the entered data as an email attachment to a specified Outlook account, streamlining data collection and reporting processes. The form is specifically an employee onboarding form for my dream company, a chess arena renting out chess equipment to customers and also offering food and beverages to customers. The form can be modified for other use cases, such as, customer inquiries/orders forwarded to managers, automated collection of responses to surveys or feedbacks, and stock updates emailed to supervisors. 

### Key Features & Functionality  

 #### 1. Data Entry Form  
   - User-Friendly Interface: Input fields for structured data collection (e.g., employee details, customer details, survey responses, order forms).  
     
   ![Formsnip](https://github.com/user-attachments/assets/dcb00ff0-45e7-4ddd-a7c7-00a0f8936212)
   
- Input Validation: The Department and Role fields are selected from data validation lists created in one worksheet.

 ![ValidSnip](https://github.com/user-attachments/assets/6ad44fce-4b8b-4c68-a6c6-6517f834e35f)

 #### 2. Automated Email Submission (VBA Macro)  
   - Attaches Worksheet: Converts the submitted data into an Excel file.

    
     ![FormDB](https://github.com/user-attachments/assets/b621ce1a-3e4b-4601-89ae-21c894767e08)

    
   - Sends via Outlook: Uses Outlook.Application to generate and dispatch emails.

       ![WSsubmit](https://github.com/user-attachments/assets/aa4ac6cf-513d-4af5-9814-50d03c160cad)

   - Customizable Email Content 
   - Error Handling: Alerts user if email fails to send.  

### Why This Project?  
 - Eliminates Manual Processes  – No more saving and emailing files manually.  
 - Reduces Errors – Structured data entry minimizes typos/missing info.
