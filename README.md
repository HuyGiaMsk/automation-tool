## **Guide to run this tool** 

Please note that all scripts in this tool are currently written for the user who primarily using Windows.

### Install the env (python, download source, virtual env):
.\script\install_env

### Console approach: 
- You must provide mandatory information in the `.\input\InvokedClasses.properties` file and its subsequent properties files which represent the settings for each running task.

vi .\input\InvokedClasses.properties .\input\ExampleTask.properties

- Perform the provided automation task by invoking:

.\start_app console

### GUI approach:
- Perform the GUI app by invoking:

.\start_app gui

- You should provide mandatory metadata for the automated task shown in the GUI.

###### After the run, all logs will be stored in `.\log` <br><br> These logs may be helpful for investigating any issues we faced during the running process.


