from exchangelib import Credentials, Account, DELEGATE, Configuration
from email.header import decode_header
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import mysql.connector
import re
from datetime import datetime
import customtkinter as ctk
import os
import uuid
import html2text
import win32com.client

# Function to connect to the database
def connect_to_db():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="moiz",
        database="caa"
    )

def convert_ewsdatetime_to_datetime(ewsdatetime):
    # Convert EWSDateTime to a datetime object
    return datetime(
        ewsdatetime.year,
        ewsdatetime.month,
        ewsdatetime.day,
        ewsdatetime.hour,
        ewsdatetime.minute,
        ewsdatetime.second
    )

# Function to execute a query and fetch a single result
def fetch_one(query, params=None):
    connection = connect_to_db()
    cursor = connection.cursor()
    cursor.execute(query, params)
    result = cursor.fetchone()
    connection.close()
    return result

# Function to execute a query and fetch all results
def fetch_all(query, params=None):
    connection = connect_to_db()
    cursor = connection.cursor()
    cursor.execute(query, params)
    results = cursor.fetchall()
    connection.close()
    return results
    
# Function to scrape and store into the Database
def scrape(exchange_username, exchange_password, userID):
    try:

        # Set up Exchange Server connection
        credentials = Credentials(username=f"HQCAA.NET\\{exchange_username}", password=exchange_password)        
        email = f"{exchange_username}@caapakistan.com.pk"
        print(email)
        account = Account(email, credentials=credentials, autodiscover=True)

        # Function to generate a unique filename
        def generate_unique_filename(directory, filename):
            base, ext = os.path.splitext(filename)
            unique_name = f"{base}_{uuid.uuid4().hex}{ext}"
            return unique_name
        
        conn = connect_to_db()
        cursor = conn.cursor()

        # Define the directory for storing attachments
        attachment_dir = "Attachment"  # Replace with the actual path

        for item in account.inbox.filter(is_read=False):  # You can adjust this filter as needed
            item.is_read = True  # Mark the email as read
            item.save()  # Save the changes

            subject = item.subject

            html_content = item.body
            plain_text = html2text.html2text(html_content)
            email_body = plain_text

            date_obj = convert_ewsdatetime_to_datetime(item.datetime_received)

            sender_email = item.sender.email_address

            print("\nSubject: " + subject)
            print("\nBody: " + email_body)
            print("\nDate: " + str(date_obj))
            print("\nSender: " + sender_email)

            # Insert email information into the database
            logged_in_user_id = userID 
            insert_query = "INSERT INTO email (userID, subject, body, date, senderEmail) VALUES (%s, %s, %s, %s, %s)"
            values = (logged_in_user_id, subject, email_body, date_obj, sender_email)
            cursor.execute(insert_query, values)
            conn.commit()

            select_email_id_query = "SELECT LAST_INSERT_ID()"
            cursor.execute(select_email_id_query)
            email_id = cursor.fetchone()[0]

            # Process attachments
            for attachment in item.attachments:
                attachment_size_mb = round(len(attachment.content) / (1024 * 1024), 1)

                if attachment.name:
                    unique_filename = generate_unique_filename(attachment_dir, attachment.name)
                    attachment_path = os.path.join(attachment_dir, unique_filename)

                    with open(attachment_path, "wb") as attachment_file:
                        attachment_file.write(attachment.content)

                    print("\nAttachment Name: " + unique_filename)
                    print("\nAttachment Path: " + attachment_path)
                    print("\nAttachment Size: " + str(attachment_size_mb))
                    print("\nLogged-in User ID: " + str(logged_in_user_id))

                    # Insert attachment information into the database
                    attachment_query = "INSERT INTO attachment (emailID, userID, AttachmentName, AttachmentSize, AttachmentPath) VALUES (%s, %s, %s, %s, %s)"
                    attachment_values = (email_id, logged_in_user_id, unique_filename, attachment_size_mb, attachment_path)
                    cursor.execute(attachment_query, attachment_values)
                    conn.commit() 
    
        conn.close()
    
    except Exception as e:
        # Handle exceptions here (e.g., display an error message)
        print("Error:", e)

def share(link):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Create a new mail item
        mail = outlook.CreateItem(0)
        mail.Body = link

        # Display the email form
        mail.Display()

    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Function to display emails and their attachments for a given userID
def display_emails_and_attachments(user_id):
    try:
        # Connect to the MySQL database
        connection = connect_to_db()
        cursor = connection.cursor()

        # Define the SQL query to retrieve emails and their attachments
        query = """
            SELECT Email.EmailID, Email.subject, Email.body, Email.date, Email.senderEmail,
                   Attachment.AttachmentName, Attachment.AttachmentSize, Attachment.AttachmentPath
            FROM Email
            LEFT JOIN Attachment ON Email.EmailID = Attachment.EmailID
            WHERE Email.userID = %s
        """

        # Execute the query with the specified user_id
        cursor.execute(query, (user_id,))

        # Fetch all matching records
        results = cursor.fetchall()

        records_window = ctk.CTkToplevel(root)
        records_window.geometry("700x550")
        records_window.attributes("-topmost", True)

        records_window.title("Records")

        style = ttk.Style()
        style.configure("Treeview", background="light blue", foreground="black", fieldbackground="dark gray")
        
        # Create a treeview widget to display the results
        tree = ttk.Treeview(records_window, style="Treeview")
        tree["columns"] = ("Email_ID","Subject", "Body", "Date", "Sender Email", "Attachment Name", "Attachment Size", "Attachment Path")

        # Define column headings
        tree.heading("#0", text="Email_ID")
        tree.heading("#1", text="Subject")
        tree.heading("#2", text="Body")
        tree.heading("#3", text="Date")
        tree.heading("#4", text="Sender Email")
        tree.heading("#5", text="Attachment Name")
        tree.heading("#6", text="Attachment Size")
        tree.heading("#7", text="Attachment Path")

        # Populate the treeview with data
        for row in results:
            email_id, subject, body, date, sender_email, attachment_name, attachment_size, attachment_path = row

            tree.insert(
                "", "end",
                text=email_id,  # Email_ID
                values=(subject, body, date, sender_email, attachment_name, attachment_size, attachment_path)
            )

            # Function to open the attachment using the attachment_path
            def open_attachment(event):
                selected_item = tree.selection()
                if selected_item:
                    item = tree.item(selected_item)
                    attachment_path = item['values'][6]  # Assuming the attachment_path is in the 8th column
                    if attachment_path:
                        import os
                        os.startfile(attachment_path)  # Opens the file with the default program

            # Function to open the attachment using the attachment_path
            def share_attachment(event):
                share(attachment_path)
                        
        # Bind the open_attachment function to the Treeview
        tree.bind("<A>", open_attachment)
        tree.bind("<a>", open_attachment)
                        
        # Bind the open_attachment function to the Treeview
        tree.bind("<S>", share_attachment)
        tree.bind("<s>", share_attachment)
         
        # Pack the treeview
        tree.pack(fill="both", expand=True)

        # Close the database connection
        cursor.close()
        connection.close()

    except mysql.connector.Error as err:
        print("Error:", err)

def search(keyword, criteria):
    try:
        # Connect to the MySQL database
        connection = connect_to_db()
        cursor = connection.cursor()

        # Define the SQL query based on the search criteria
        if criteria == "keyword":
            search_query = """
                SELECT Email.EmailID, Email.subject, Email.body, Email.date, Email.senderEmail,
                Attachment.AttachmentName, Attachment.AttachmentSize, Attachment.AttachmentPath
                FROM Email
                LEFT JOIN Attachment ON Email.EmailID = Attachment.EmailID
                WHERE Email.subject LIKE %s OR Email.body LIKE %s
            """
            # Execute the query with the keyword as a wildcard search
            keyword_search = f"%{keyword}%"
            cursor.execute(search_query, (keyword_search, keyword_search))

        elif criteria == "date":
            search_query = """
                SELECT Email.EmailID, Email.subject, Email.body, Email.date, Email.senderEmail,
                Attachment.AttachmentName, Attachment.AttachmentSize, Attachment.AttachmentPath
                FROM Email
                LEFT JOIN Attachment ON Email.EmailID = Attachment.EmailID
                WHERE Email.date LIKE %s
            """

            # Execute the query with the keyword as a wildcard search
            keyword_search = f"%{keyword}%"
            cursor.execute(search_query, (keyword_search,))

        elif criteria == "sender":
            search_query = """
                SELECT Email.EmailID, Email.subject, Email.body, Email.date, Email.senderEmail,
                Attachment.AttachmentName, Attachment.AttachmentSize, Attachment.AttachmentPath
                FROM Email
                LEFT JOIN Attachment ON Email.EmailID = Attachment.EmailID
                WHERE Email.senderEmail LIKE %s
            """
            # Execute the query with the keyword as a wildcard search
            keyword_search = f"%{keyword}%"
            cursor.execute(search_query, (keyword_search,))

        # Fetch all matching records
        search_results = cursor.fetchall()

        # Print the results (you can remove this section if not needed)
        for row in search_results:
            email_id, subject, body, date, sender_email, attachment_name, attachment_size, attachment_path = row
            print(f"Email ID: {email_id}")
            print(f"Subject: {subject}")
            print(f"Body: {body}")
            print(f"Date: {date}")
            print(f"Sender Email: {sender_email}")
            if attachment_name:
                print(f"Attachment Name: {attachment_name}")
                print(f"Attachment Size: {attachment_size} MB")
                print(f"Attachment Path: {attachment_path}")
            print("-" * 30)

        # Create and configure the search window
        search_window = tk.Tk()
        search_window.geometry("700x550")
        search_window.attributes("-topmost", True)
        search_window.title("Search Results")

        style = ttk.Style()
        style.configure("Treeview", background="light blue", foreground="black", fieldbackground="dark gray")

        # Create a treeview widget to display the search results
        tree = ttk.Treeview(search_window, style="Treeview")
        tree["columns"] = ("Email_ID", "Subject", "Body", "Date", "Sender Email", "Attachment Name", "Attachment Size", "Attachment Path")

        # Define column headings
        tree.heading("#0", text="Email_ID")
        tree.heading("#1", text="Subject")
        tree.heading("#2", text="Body")
        tree.heading("#3", text="Date")
        tree.heading("#4", text="Sender Email")
        tree.heading("#5", text="Attachment Name")
        tree.heading("#6", text="Attachment Size")
        tree.heading("#7", text="Attachment Path")

        # Populate the treeview with search results
        for row in search_results:
            email_id, subject, body, date, sender_email, attachment_name, attachment_size, attachment_path = row
            tree.insert(
                "", "end",
                text=email_id,  # Email_ID
                values=(subject, body, date, sender_email, attachment_name, attachment_size, attachment_path)
            )
        
        # Function to open the attachment using the attachment_path
        def open_attachment(event):
            selected_item = tree.selection()
            if selected_item:
                item = tree.item(selected_item)
                attachment_path = item['values'][6]  # Assuming the attachment_path is in the 8th column
                if attachment_path:
                    import os
                    os.startfile(attachment_path)  # Opens the file with the default program
                        
        # Bind the open_attachment function to the Treeview
        tree.bind("<A>", open_attachment)
        tree.bind("<a>", open_attachment)         

        # Pack the treeview
        tree.pack(fill="both", expand=True)

        # Close the database connection
        cursor.close()
        connection.close()

        # Start the main loop of the search window
        search_window.mainloop()

        return search_results

    except mysql.connector.Error as err:
        print("Error:", err)
        return None  # Return None in case of an error

def open_search_window():
    global search_window
    search_window = ctk.CTkToplevel()
    search_window.geometry("700x550")
    search_window.attributes("-topmost", True)

    frame = ctk.CTkFrame(master=search_window)
    frame.pack(pady=20, padx=60, fill="both", expand=True)

    search_window.title("Search Records")

    search_window_label = ctk.CTkLabel(master=frame, text="Search Records", font=("Roboto", 28, "bold"))
    search_window_label.pack(pady=12, padx = 8)

    search_by_subject = ctk.CTkEntry(master=frame, placeholder_text="Search by Keyword")
    search_by_subject.pack(pady=8, padx=10)

    search_button = ctk.CTkButton(master=frame, text="Search", command=lambda: search(search_by_subject.get(),"keyword"))
    search_button.pack(pady=8, padx=10)

    search_by_date = ctk.CTkEntry(master=frame, placeholder_text="Search by Date")
    search_by_date.pack(pady=8, padx=10)

    search_button2 = ctk.CTkButton(master=frame, text="Search", command=lambda: search(search_by_date.get(),"date"))
    search_button2.pack(pady=8, padx=10)

    search_by_sender = ctk.CTkEntry(master=frame, placeholder_text="Search by Sender")
    search_by_sender.pack(pady=8, padx=10)

    search_button3 = ctk.CTkButton(master=frame, text="Search", command=lambda: search(search_by_sender.get(),"sender"))
    search_button3.pack(pady=8, padx=10)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

def get_user_id_by_username(username):
    try:
        conn = connect_to_db()
        cursor = conn.cursor()

        # Define the SQL query to select the user's ID based on the username
        query = "SELECT userID FROM user WHERE username = %s"

        # Execute the query with the provided username as a parameter
        cursor.execute(query, (username,))

        # Fetch the result
        result = cursor.fetchone()

        if result:
            return result[0]  # Return the user ID
        else:
            return None  # User not found

    except mysql.connector.Error as error:
        print(f"Error: {error}")
    finally:
        if 'connection' in locals() and conn.is_connected():
            cursor.close()
            conn.close()

# Root Email Configuration Window
root = ctk.CTk()
root.title("Login")

# Set the window size
root.geometry("700x550")

frame = ctk.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

root.title("Email Configuration Form")

# Create and configure labels
label2 = ctk.CTkLabel(master=frame, text="Configure Email", font=("Roboto", 28, "bold"))
label2.pack(pady=12, padx=10)

subtitle_label = ctk.CTkLabel(master=frame, text="Email Configuration")
subtitle_label.pack(pady=8, padx= 10)

# Create entry widgets
entry_exchange_username = ctk.CTkEntry(master=frame, placeholder_text="Exchange Username")
entry_exchange_username.pack(pady=8, padx=10)

entry_exchange_password = ctk.CTkEntry(master=frame, placeholder_text="Exchange Password", show="*")
entry_exchange_password.pack(pady=8, padx=10)

def success2():
    result = scrape(entry_exchange_username.get(), entry_exchange_password.get(), get_user_id_by_username(entry_exchange_username.get()))
    if result:    
        success_label = ctk.CTkLabel(master=frame, text="Extraction Successful")
        success_label.pack(pady=8, padx= 10)
    else:
        success_label = ctk.CTkLabel(master=frame, text="Extraction Successful")
        success_label.pack(pady=8, padx= 10)

# Create Submit button
submit_button = ctk.CTkButton(master=frame, text="Submit", command=lambda: success2())
submit_button.pack(pady=15, padx=10)
    
records_button = ctk.CTkButton(master=frame, text="Show All Records", command=lambda: display_emails_and_attachments(get_user_id_by_username(entry_exchange_username.get())))
records_button.pack(pady=15, padx=10)

search_window = ctk.CTkButton(master=frame, text="Search Records", command=lambda: open_search_window())
search_window.pack(pady=15, padx=10)


# Start the tkinter main loop
root.mainloop()