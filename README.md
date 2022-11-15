# Afas email connector

[Python](http://python.org) tool that connects [Afas](https://afas.com) with an email server, written for an internship in June 2022.

## Development

Requirements:

- [Python](https://python.org)
- [pip](https://pypi.org/project/pip)

```bash
git clone https://github.com/Steveplays28/afas-email-connector.git
cd afas-email-connector

pip install python-dotenv
pip install requests

python src/afas_email_connector.py
```

A VSCode run configuration is also included.

### Environment variables

Create a `secrets.env` file in [`src/`](https://github.com/Steveplays28/afas-email-connector/tree/main/src) with the following content:

```env
# Email account credentials
USERNAME = "address@domain"
PASSWORD = "password"

AFAS_UPDATECONNECTOR_API_TOKEN = "<token><version>N</version><data>TOKEN</data></token>"
```

Email server is currently hardcoded to be [`outlook.office365.com` (Outlook)](https://outlook.office365.com).

## Notes

Some important info about the Afas environment that helped me make this.

- `debiteurnummer` = `999999`
- `dossier item type` = `21`

## License

This project is licensed under the [MIT license](https://mit-license.org), see [LICENSE](https://github.com/Steveplays28/afas-email-connector/blob/main/LICENSE).
