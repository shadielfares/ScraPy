# ScraPy ☕📧

ScraPy is a Python-based tool designed to send out research emails efficiently. This repository contains scripts that automate the process of scraping data and sending emails. Instead of sending specific emails to each person, I've opted for a more streamlined approach by asking for coffee chats, believing it would be more effective. In the future, I intend to implement an AI agent that will scrape each professor's entry and tailor the emails specifically to them. 🤖✨

NOTE: Clicking on the "Wiki" section at the top of this repo will show you some of the results this project yielded.

## Prerequisites ✅

Before you can run any scripts in this project, make sure you have the latest version of [Python 3](https://www.python.org/downloads/) installed.

1. **Install Python 3**: Follow the links above to download and install the latest versions.
  
2. **Create a Virtual Environment**:
   - Navigate to your project directory:
     ```bash
     cd path/to/your/project
     ```
   - Create a virtual environment:
     ```bash
     python -m venv .venv
     ```
   - Activate the virtual environment:
     - **On Windows**: 
       - Open PowerShell as Administrator and run:
         ```powershell
         Set-ExecutionPolicy Unrestricted
         ```
       - Then activate the virtual environment:
         ```powershell
         .\.venv\Scripts\activate
         ```

     - **On macOS/Linux**:
       - Activate the virtual environment using:
         ```bash
         source .venv/bin/activate
         ```

3. **Clone the Repository Inside the `.venv` Directory**:
   - After activating your virtual environment, clone the ScraPy repository:
   - Navigate into the cloned repository:
     ```bash
     cd ScraPy
     ```

4. **Install Required Packages**: You can install the required packages using pip:
   ```bash
   pip install -r requirements.txt
   ```

Here are the required packages listed in `requirements.txt`:

```
beautifulsoup4==4.12.3
certifi==2024.6.2
charset-normalizer==3.3.2
et-xmlfile==1.1.0
idna==3.7
openpyxl==3.1.3
python-dotenv==1.0.1
requests==2.32.3
soupsieve==2.5
urllib3==2.2.1
```

## Important File Setup 📋

- **Replace `DUMMYDOC.xlsx`**: Ensure you replace `DUMMYDOC.xlsx` with an empty Excel workbook, replacing its current existence in the project folder.

## Demo Run-through 🚀

Here's a step-by-step guide to running a script from the `scraperMagic` folder:

1. **Navigate to the Project Directory**:
   ```bash
   cd scraperMagic
   ```

2. **Run a Script**: Assuming you have a script named `example_script.py` in the `scraperMagic` folder, you can run it as follows:
   ```bash
   python example_script.py
   ```

### Example: Running `emailer.py` ✉️

If you have a script named `send_emails.py` in the `scraperMagic` folder, you can run it to send out research emails. Make sure to configure any necessary environment variables or input files as required by the script.

- **Configure Gmail for `emailer.py`**:
  To send emails from a personal Gmail address, you'll need to create an app password. Please read the instructions in the `emailer.py` script to set this up correctly.

- **Attaching Files**:
  If you want to attach any files to the email, read through the corresponding code block in `emailer.py` and follow the comments for instructions on how to do this.

```bash
python emailer.py
```

## Disclaimer ⚠️

I am not liable for any damage, spam, or liabilities this project may cause. This includes but is not limited to any unintended consequences of using this tool, such as your email being flagged as spam or any impact on your reputation. 

I have chosen not to include extensive documentation exploring all features and scripts, as the level of abstraction was necessary to limit its usability to those with a specific purpose who have the capabilities to modify the script and use it for good. By using this tool, you acknowledge that you understand the risks involved and agree to take full responsibility for any actions taken as a result of using this project.

## License 📄

This project is licensed under the CCC License - see the LICENSE file for details.

## Contributing 🤝

Feel free to fork this repository and submit pull requests. For major changes, please open an issue first to discuss what you would like to change.
