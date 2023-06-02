# Salesforce Data Model to Excel

What this script does:

- retrieve sfdx project to the folder
- gets all **Custom** **Fields**, **Standard** **Fields** and **Standard** **Value** **Sets** (Picklists)
- **full** **global** **picklist** value set support, standard picklists duplicates covered
- **standard** **lookups** support
- parse them to excel and create data model for each object in separate sheet
- creates 'All' sheet with whole data model in one sheet
- **sheets** and **rows** sorted **alphabetically** (but not by ASCII values, true alphabetically)

What can be configured (config.js):
- org authorization (username, password, security_token, client_id, client_secret)
- which standard value sets you want in which sheet (standardValueSets)
- auth_url (auth_url)
- which objects are interesting for you (object_list)

TODO
- type and label for standard fields that are not picklists

Benefits:
You can filter the excel.
Example: check all datetime fields in the system for selected objects.

<img width="1436" alt="image" src="https://github.com/Richard98PL/salesforce-data-model/assets/41301282/36760a23-2132-4c80-bef1-d84ce8431aea">

# Before
![image](https://github.com/Richard98PL/salesforce-data-model/assets/41301282/b49d8eab-4131-4c42-8d22-0a419db97a61)

# After
![image](https://github.com/Richard98PL/salesforce-data-model/assets/41301282/734d5a3b-50fe-44fa-b84e-58b59f9f53c6)

# Filtering
<img width="1383" alt="image" src="https://github.com/Richard98PL/salesforce-data-model/assets/41301282/2bf9392e-93d7-4746-9935-4d192da81170">
