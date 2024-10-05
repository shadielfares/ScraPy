# ScraPy

ScraPy is a Python-based tool designed to send out research emails efficiently. This repository contains scripts that automate the process of scraping data and sending emails.

## Prerequisites

Before you can run any scripts in this project, make sure you have the following dependencies installed. You can install them using `pip`:

```bash
pip install -r requirements.txt
```

Here are the required packages:
- `beautifulsoup4==4.12.3`
- `certifi==2024.6.2`
- `charset-normalizer==3.3.2`
- `et-xmlfile==1.1.0`
- `idna==3.7`
- `openpyxl==3.1.3`
- `python-dotenv==1.0.1`
- `requests==2.32.3`
- `soupsieve==2.5`
- `urllib3==2.2.1`

## Demo Run-through

Here's a step-by-step guide to running a script from the `scraperMagic` folder:

1. **Navigate to the Project Directory**:
    ```bash
    cd ScraPy/scraperMagic
    ```

2. **Run a Script**:
    Assuming you have a script named `example_script.py` in the `scraperMagic` folder, you can run it as follows:
    ```bash
    python example_script.py
    ```

### Example: Running `send_emails.py`

If you have a script named `send_emails.py` in the `scraperMagic` folder, you can run it to send out research emails. Make sure to configure any necessary environment variables or input files as required by the script.

```bash
python send_emails.py
```

## Configuration

Ensure you have the necessary configuration set up, such as environment variables. If the scripts use a `.env` file, create one in the root directory of the project and add your variables there.

Example `.env` file:
```
EMAIL_HOST=smtp.example.com
EMAIL_PORT=587
EMAIL_USER=your_email@example.com
EMAIL_PASS=your_password
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Feel free to fork this repository and submit pull requests. For major changes, please open an issue first to discuss what you would like to change.

---

This README provides an overview of the ScraPy project, including prerequisites, a demo run-through, and configuration details. For more detailed information, refer to individual script documentation within the repository.
