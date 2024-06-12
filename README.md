# OFT to HTML Converter

The project is a Java application designed to convert Outlook Template (OFT) files into HTML format. 

## Prerequisites

This project requires Gradle to compile and run. On a Mac, an easy way to install Gradle is by using Homebrew. If Homebrew is not installed on your system, you can install it by executing the following command in a terminal:

```shell
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

After installing Homebrew, you can install Gradle by running:

```shell
brew install gradle
```

## Project Setup

To set up the project, navigate to the project's root directory in the terminal. Compile the project with the following Gradle command:

```shell
gradle build
```

This command compiles the project and generates the necessary output files in the `build` directory.

## Running the Project

To run the project and convert OFT files to HTML format, use the command:

```shell
gradle run
```

This command processes all `.oft` files located in the "input" directory and saves the converted `.html` files in the "output" directory.

### Input and Output Directories

- **Input Directory**: Place your `.oft` files in the "input" directory at the root of the project. The application scans this directory for `.oft` files to convert.

- **Output Directory**: The converted `.html` files are saved in the "output" directory, which is also located at the project's root. If the "output" directory does not exist, the application will create it.

### Watermark

The generated files will contain a watermark from Aspose, the company that provides the file format conversion tools. This watermark can be removed by obtaining a license for "Aspose.Email for Java".

Aspose offers free temporary licenses at https://purchase.aspose.com/temporary-license/

To apply your license, add the file "Aspose.EmailforJava.lic" to the root folder of this project before running it.