# commission-comparer-infynity

## Generating a binary file
NOTE: Make sure you generate the binary in the same **operating system** it is going to be used and ensure the **python version** of the project is the same installed in your OS.

1. Clone project from git
1. Start your development environment
1. Install all dependencies in your environment
1. Generate binary file
    1. Run `pyinstaller name_of_main_file.py --onefile` inside the project directory

       This will create two new directories inside your project `build` and `dist`.

    1. Your cli binary will be inside the `dist` directory

       You can rename the cli file to anything you want.

       Don't forget to make it an executable by running `chmod +x [name_of_file]`.
1. Now try running `python cli.py --help` to have a list of commands.

> DEV NOTES: It will not work because the path to the generated files is pointing to inside the project at the moment.


==========
Create the following Directories in inputs directory: 
Qaisars-Mac-mini:inputs qaisar$ l
total 16
drwxr-xr-x@  8 qaisar  staff   256B 30 Jul 00:32 .
drwxr-xr-x@ 19 qaisar  staff   608B 30 Jul 20:15 ..
-rw-r--r--@  1 qaisar  staff   6.0K 28 Jul 13:31 .DS_Store
drwxr-xr-x@  4 qaisar  staff   128B 30 Jul 00:32 downloads
drwxr-xr-x@  4 qaisar  staff   128B 30 Jul 00:28 infynity
drwxr-xr-x@  3 qaisar  staff    96B 30 Jul 00:32 loankit
drwxr-xr-x@  8 qaisar  staff   256B 30 Jul 00:16 referrer_list
drwxr-xr-x@  4 qaisar  staff   128B 30 Jul 00:08 test_files

Create Outputs directory
