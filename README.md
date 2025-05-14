# Automated Data Entry Form with Email Submission (Excel VBA)

### Project Overview

The project is an Excel-based data entry form with a VBA-powered "Submit" button that automatically sends the entered data as an email attachment to a specified Outlook account, streamlining data collection and reporting processes. 

### Use Cases

The form is specifically an employee onboarding form for my dream company, a chess arena renting out chess equipment to customers and also offering food and beverages to customers. The form can be modified for other use cases, such as, customer inquiries/orders forwarded to managers, automated collection of responses to surveys or feedbacks, and stock updates emailed to supervisors. 

### Key Features & Functionality  

 #### 1. Data Entry Form  
   - User-Friendly Interface: Input fields for structured data collection (e.g., employee details, customer details, survey responses, order forms).  
     
   ![Formsnip](https://github.com/user-attachments/assets/dcb00ff0-45e7-4ddd-a7c7-00a0f8936212)
   
- Input Validation: The Department and Role fields are selected from data validation lists created in one worksheet. Entries in the data validation lists are derived by asking AI chatbot (DeepSeek) 

 ![ValidSnip](https://github.com/user-attachments/assets/6ad44fce-4b8b-4c68-a6c6-6517f834e35f)

 #### 2. Automated Email Submission (VBA Macro)  
   - Attaches Worksheet: Converts the submitted data into an Excel file.

    
     ![FormDB](https://github.com/user-attachments/assets/b621ce1a-3e4b-4601-89ae-21c894767e08)

    
   - Sends via Outlook: Uses Outlook.Application to generate and dispatch emails.

       ![WSsubmit](https://github.com/user-attachments/assets/aa4ac6cf-513d-4af5-9814-50d03c160cad)

   - Customizable Email Content 
   - Error Handling: Alerts user if email fails to send.  

### Tools and Technologies

#### **1. Microsoft Excel**
- **Primary platform** for data entry forms and storage
- **Worksheets** for structured data organization
- **Tables** for managing submitted records
- **Form Controls** (buttons, input fields, labels)

#### **2. VBA (Visual Basic for Applications)**
- **Automation backbone** for all functionality
- **Key uses**:
  - Form validation (data format checks)
  - Email generation and sending
  - File attachment handling
  - Error handling and user notifications

#### **3. Microsoft Outlook for email delivery
#### **4. AI chatbot (DeepSeek) for creating the data validation lists, and editing the VBA macros



