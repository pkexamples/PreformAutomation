# Import PKSLRemote module and initialize it
from PKSLRemote import Common, PK2650
Common.init(RunForm, Executor)
# Set preform ID, diameter, template; load preform and locate tip
PK2650.SetPreformId('Test', 18.25)
PK2650.SetTemplate('LM3317.pftm')
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