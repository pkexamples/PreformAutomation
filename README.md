# PK 2650 / 2660 Preform Measurement Automation with PKSL Remote

Due to the complexity and long duration of preform measurements, it is typically desirable to automate measurements to allow for dynamic control over measurements based upon real-time intermediate measurement results, rather than to rely on fully pre-programmed measurement sequences that statically define all measurement locations and parameters. In a typical production environment, an external program controls the measurement workflow, processes results, and may interact with factory data services. A PKSL Remote remote control client program can be written in any .NET Framework language or any language that supports COM (Component Object Model).

This guide consists of this README summary along with several example scripts that utilize the `PKSLRemote` script package (see the `Scripts` directory), several example script-based results processors (see the `Utilities` directory), and an example measurement sequence file for the example Sequence File remote client application (in `SequenceFile`).

Please contact [PKSL Support](mailto:support@pkinetics.com) to acquire the latest PKSL and PKSL Remote software installers. PKSL Remote can be installed on any Windows 7 or later computer, and its install program places a desktop shortcut to the locally installed documentation.

## PKSL Remote

The PKSL Remote interface consists of a RemoteHost.py script and a package of scripts called the `PKSLRemote` package. The `PKSLRemote` script package consists of a collection of modules, each containing functions that define the PKSL Remote interface commands. There is a `Common` command module, as well as a module for each PSKL instrument type. When executed, the RemoteHost.py script places a control in the central "runtime" panel of the PKSL Main Application's Measurement Mode that displays commands received and replies sent to a remote control program via a simple communication service. The script then starts the communication service and waits for a a remote control program to connect to the service and send a message. It then processes the message, and replies with a result. This message processing loop continues until the `Abort` button is clicked.  These components are all installed with the PKSL Remote installation program along with some example remote control programs (with source code) and the PKSLRemote package documentation, which lists the available commands for each module.

## Measurement result processing

The `MarkSlice` command in both the `PK2650` and `PK2660` modules allows you to specify a results processor (explained below) to execute after all the angular position scans have been completed for a specific z-postion. The `MarkSlice` command should be issued after issuing the `NewSlice` command. After issuing the sequence of commands needed to execute the scan(s) and compute the slice result, the `ProcessResults` command causes the specified results processor(s) to execute. See the illustrative snippet below:

```py
for z in z_positions:
    PK2650.MoveToZ(z)
    PK2650.NewSlice()
    # mark the next "slice" to be measured to be processed by
    # results processors with Service Names Deflection_ASCII
    # and Power_ASCII
    PK2650.MarkSlice("Deflection_ASCII")
    PK2650.MarkSlice("Power_ASCII")
    # measure the slice (w_positions is a list containing one
    # or more angualar positions in degrees)
    for w in w_positions:
        PK2650.RotateToW(w)
        PK2650.LocatePreform()
        PK2650.AutoFocus()
        PK2650.HomeMotors()
        PK2650.ComputeOilIndex()
        PK2650.Measure()
    # all angular position scans have completed; compute
    # the results and process them with the specified results
    # processors
    PK2650.ComputeSlice()
    Common.ProcessResults()
```

A results processor often is employed to persist in-memory measurement results to a text file so that an external program (such as a remote control program) can read it. Results processors could also perform other operations on the measurement result data, such as running custom algorithms on raw data, comparing results to product specifications, and/or creating reports.

The `PKSummaryFile.py` file that is installed into the "Utilities" sub-directory of the PKSL "Scripts" directory with the PKSL software is an example of a script-based results processor, but others can be added that follow the same pattern. All results processors must have comments at the top to give them a "Service Name" and a "Friendly Name". Results processors can also be written in a .NET Framework language using the Managed Extensibility Framework (MEF). See the PKSL User's Guide for more details on MEF results processors.

## Automating with IronPython scripts

The `PKSLRemote` package may also be utilized as an IronPython package, so it is possible to implement automation without writing a remote control program. This is also a good way to prototype and test the behavior you desire a remote control program to have if you choose that automation path. To do this, you need to import the `Common` module and either the `PK2650` or `PK2660`, depending on your instrument type, and then call the  `Common.init` function, passing in the `RunForm` and `Executor` variables that are inserted automatically into all executed PSKL scripts.

```py
from PKSLRemote import Common, PK2650
Common.init(RunForm, Executor)
```

The following is the full listing for the `PK2650_ThreeSlices_wFocusRetry.py` example script found in the `Scripts` directory.

```py
# Import PKSLRemote module and initialize it
from PKSLRemote import Common, PK2650
Common.init(RunForm, Executor)

# Set preform ID, diameter, template; load preform and locate tip
PK2650.SetPreformId('Test', 18.25)
PK2650.SetTemplate('MyTemplate.pftm')
PK2650.MoveZManual('Load preform and position at tank opening')
PK2650.LocateTip()

# define w and z locations in degrees and mm
num_angles = 5
w_positions = [x * 360 / num_angles for x in range(num_angles)]
z_positions = (60, 50, 40)
max_focus_attempts = 3

# immerse preform and equilibrate
PK2650.MoveToZ(max(z_positions))
PK2650.EquilibratePreform()
for z in z_positions:
    PK2650.MoveToZ(z)
    PK2650.NewSlice()
    PK2650.MarkSlice("Deflection_ASCII")
    PK2650.MarkSlice("Power_ASCII")
    for w in w_positions:
        PK2650.RotateToW(w)
        # try to focus up to max_focus_attempts before giving up
        focus_count = 0
        while True:
            focus_count += 1
            try:
                PK2650.LocatePreform()
                PK2650.AutoFocus()
                break
            except:
                if focus_count == max_focus_attempts:
                    raise
        PK2650.HomeMotors()
        PK2650.ComputeOilIndex()
        PK2650.Measure()
    PK2650.ComputeSlice()
    Common.ProcessResults()
PK2650.MoveToZ(0)
Common.Pause()
```

## Appendix: Utilizing your existing code in Python scripts

If you have your own .NET Framework libraries that you would like to use in your automation script, they may be placed in the `Utilities` subdirectory of the PKSL `Scripts` directory, and then referenced by using the clr.AddReference() method.

```py
# Example of adding a reference to a library named FactoryDataServices.dll
clr.AddReference('FactoryDataServices')
```

Once the reference is added, namespaces and/or class names can be imported.

```py
# Example of importing a class named "DatabaseService" from the "FactoryDataServices" namespace.
from FactoryDataServices import DatabaseService
```

If you have existing or 3rd party Component Object Model (COM) libraries you would like to use, they too can be utilized by using the .NET Framework's `System.Activator`.

```py
from System import Activator, Type
xlType = Type.GetTypeFromProgID("Excel.Application")
excel = Activator.CreateInstance(xlType)
wb = excel.Workbooks.Add()
excel.Quit()
```
