# OutlookUtils
A command-line tool for invoking a local Outlook client based on Python.

# Usage Example
```Shell
# Get Email Subject
outlook_utils.py get-emails-subject --email-addr <email_address>
# Get Email Summary
outlook_utils.py get-emails-summary --email-addr <email_address>
# Send Email from local Outlook Client
outlook_utils.py send-email --to-addr <email_address> [--subject <string> --content <string>]
# Help Information
outlook_utils.py --help
outlook_utils.py get-emails-subject --help
outlook_utils.py get-emails-summary --help
outlook_utils.py send-email --help
```