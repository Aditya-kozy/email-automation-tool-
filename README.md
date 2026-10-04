# Email Automation Tool

A Python tool to send personalized payment reminder emails to multiple contacts from an Excel file.

## What This Tool Does

- Reads names and email addresses from `contacts.xlsx`  
- Sends a personalized "Payment Reminder" email to each person  
- Prints success/failure messages in the terminal  
- Continues to the next contact if one fails

## Project Files

| File | Description |
| :---- | :---- |
| `email_sender.py` | The main Python script |
| `contacts.xlsx` | Sample Excel file with 3 test contacts |
| `requirements.txt` | Python dependencies to install |
| `README.md` | This file with setup instructions |

&nbsp;

---

## Step 1: Install Python Libraries

Open your terminal/command prompt in this folder and run:

```sh
pip install -r requirements.txt
```

&nbsp;

Or manually install:

```sh
pip install openpyxl
```

&nbsp;

---

## Step 2: Set Up Gmail App Password

**IMPORTANT:** You cannot use your regular Gmail password. You must create an App Password.

### Why Use an App Password?

- Gmail requires "App Passwords" for third-party applications  
- This is a 16-character password that gives access only to email sending  
- It's more secure than using your main password

### Step-by-Step Instructions:

1. **Enable 2-Step Verification** (required for App Passwords)

   - Go to [https\://myaccount.google.com/](https://myaccount.google.com/)  
   - Click "Security" in the left menu  
   - Under "Signing in to Google", click "2-Step Verification"  
   - Click "Get started" and follow the prompts  
   - Complete the setup (you'll need your phone)  
2. **Generate an App Password**

   - After enabling 2-Step Verification, go back to Security settings  
   - Under "Signing in to Google", click "App passwords"  
   - You may need to sign in again  
   - At the bottom, click "Select app" → choose "Mail"  
   - Click "Select device" → choose "Other (Custom name)"  
   - Type: `Email Automation Tool`  
   - Click "Generate"  
   - **COPY THE 16-CHARACTER PASSWORD** (it looks like: `abcd efgh ijkl mnop`)  
   - Click "Done"  
3. **Save Your App Password**

   - The password will ONLY be shown once  
   - Save it in a secure place (password manager)  
   - You'll need it every time you run this script

---

## Step 3: Run the Tool

1. Make sure `contacts.xlsx` is in the same folder as `email_sender.py`

2. Run the script:

```sh
python email_sender.py
```

3. When prompted:

   - Enter your Gmail address (e.g., `youremail@gmail.com`)  
   - Enter your App Password (16 characters, no spaces)  
4. The tool will:

   - Connect to Gmail  
   - Send emails to each contact  
   - Print success/error messages

---

## How to Customize

### Edit the Email Message

Open `email_sender.py` and find the `create_email_message` function (around line 67):

```py
email_body = f"""Hi {recipient_name},

This is a friendly reminder that your payment is due.
Please make the payment at your earliest convenience.

Thank you."""
```

&nbsp;

Change the text between the triple quotes to your custom message.

### Add More Contacts

Open `contacts.xlsx` in Excel and add rows:

- **Column A**: Name  
- **Column B**: Email

Save the file and run the script again.

### Change the Subject Line

In `email_sender.py`, find line 65:

```py
message["Subject"] = "Payment Reminder"
```

&nbsp;

Change `"Payment Reminder"` to your desired subject.

---

## Troubleshooting

### "Authentication failed" Error

- Make sure you're using an **App Password**, not your regular Gmail password  
- App Passwords require 2-Step Verification to be enabled  
- Double-check there are no spaces in the password you enter

### "File not found" Error

- Make sure `contacts.xlsx` is in the same folder as `email_sender.py`  
- Check the filename is spelled correctly (case matters on some systems)

### "No module named 'openpyxl'" Error

- Run: `pip install openpyxl`  
- If you have multiple Python versions, use: `pip3 install openpyxl`

### "Connection refused" or "Network error"

- Check your internet connection  
- Make sure your firewall isn't blocking Python  
- Try on a different network

---

