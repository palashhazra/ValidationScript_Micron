# MicronPOMigration

This project is stored in Micron folder. This project is used to validate different assets based on the data provided by Micron.

# PO Validation

Micron will provide a JSON file containing PO (Purchase Order) related data and that content should be pasted in MicronPOMigrationData.json file. Then those POs need to be validated by calling Get PO API using individual
PO ID and associated partner details stored in those individual objects of given JSON file.
After validating data from Cosmos DB and their associated given data in JSON file, we need to paste in a sheet (PO Validation result) of an excel file (DataMigrationOutput.xlsx) and need to compare those two sets of data in separate column in that same sheet mentioned previously.

## Requirements

For development, you will only need Node.js and some packages, installed in your environement.

### Node
- #### Node installation on Windows

  Just go on official Node.js website(https://nodejs.org/) and download the installer of lastest version of LTS (Long Term Support). Then follow the steps mentioned there.

If the installation was successful, you should be able to run the following command in command prompt.

    node --version or node -v
    v18.13.0

    npm --version or npm -v
    8.19.3

If you need to update `npm`, you can make it using `npm`. After running the following command, just open again the command line interface.

    npm install npm -g

## Configure app

Open folder path in VS code `C:\Users\v-pahazra\Desktop\repos\Automation Code\Micron` then run below command in Terminal.

    npm init

Then provide the details and package.json file will be created.

## Running the project

First in console, need to get inside the Micron folder. Then run below command.
    node MicronPOMigration.js

