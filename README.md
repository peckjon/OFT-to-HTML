# OFT to Word Converter

The project is a Java application designed to convert Outlook Template (OFT) files into Word (DOCX) format.

## Prerequisites

This project requires Gradle to compile and run. On a Mac, an easy way to get Gradle is by installing it with Homebrew. If you do not have Homebrew installed, you can paste the following command in a terminal:

```shell
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

To install Gradle using Homebrew, run the following command in your terminal:

```shell
brew install gradle
```

## Project Setup

Navigate to the project's root directory in the terminal. Compile the project using the following Gradle command:

```shell
gradle build
```

This command compiles your project and generates the necessary output files in the `build` directory.

## Running the Project

To run the project and convert OFT files to Word documents, use the following command:

```shell
gradle run
```
This command processes all `.oft` files located in the `input` directory and saves the converted `.docx` files in the `output` directory.

### Input and Output Directories

- **Input Directory**: Place your `.oft` files in the `input` directory at the root of the project. The application scans this directory for `.oft` files to convert.

- **Output Directory**: The converted `.docx` files are saved in the `output` directory, also located at the project's root. If the `output` directory does not exist, the application will attempt to create it.

### Watermark

The generated files will contain a watermark from Aspose, the company which provides the file format tools. It can be removed by obtaining a license for "Aspose.Email for Java".

They provide free temporary licenses at https://purchase.aspose.com/temporary-license/

To use your license, simply add the file "Aspose.EmailforJava.lic" to the root folder of this project before running it.