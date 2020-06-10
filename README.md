# Certi
>Python scripts for creating Event certificates and mailing from excel files

#Steps to run the project locally
>clone the repository

```sh
pip install -r requirement.txt
```
Edit Certi.py and add input file location or use the existing xlsx file. Add template and give image coordinates to write.

```sh
pip certi.py
```
Add your mail ID and run certi_send to send certificates to mails that are in excel file.
```sh
pip certi_send.py
```
## Release History
* 0.0.1
    * Initial Release

## Contributing

1. Fork it (https://github.com/Adarshreddyash/Certi.git)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

