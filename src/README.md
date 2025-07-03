# My Python App

This project is a Python application that can be packaged into an executable file using PyInstaller. Below are the instructions on how to run the application and build the executable.

## Project Structure

```
my-python-app
├── src
│   └── main.py          # Main Python code for the application
├── requirements.txt      # Dependencies required for the project
├── .github
│   └── workflows
│       └── build-exe.yml # GitHub Actions workflow for building the executable
└── README.md             # Project documentation
```

## Requirements

To run this application, you need to have Python installed along with the dependencies listed in `requirements.txt`. You can install the required packages using the following command:

```
pip install -r requirements.txt
```

## Running the Application

To run the application, execute the following command:

```
python src/main.py
```

## Building the Executable

To build the executable file, you can use the provided GitHub Actions workflow. This workflow will automatically set up the environment, install the necessary dependencies, and run PyInstaller to create the `.exe` file.

1. Push your changes to the `main` branch.
2. The workflow will trigger and build the executable.
3. You can find the generated `.exe` file in the workflow artifacts.

## License

This project is licensed under the MIT License.