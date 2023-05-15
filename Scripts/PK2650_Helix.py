# Import PKSLRemote module and initialize it
from PKSLRemote import Common, PK2650
Common.init(RunForm, Executor)

def Focus(max_tries):
    """Attempt to focus up to max_tries before raising error."""
    tries = 0
    while True:
        try:
            PK2650.LocatePreform()
            PK2650.AutoFocus()
            break
        except:
            tries += 1
            if tries == max_tries:
                raise

z_positions = (60, 40, 30)
w_positions = [x * 360/3 for x in range(3)]
PK2650.SetPreformId('Test', 18.25)
PK2650.MoveZManual('Load the preform')
PK2650.LocateTip()
PK2650.MoveToZ(max(z_positions))
PK2650.EquilibratePreform(5)
PK2650.NewHelix()
w_iter = iter(w_positions)
for z in z_positions:
    PK2650.MoveToZ(z)
    PK2650.RotateToW(next(w_iter))
    Focus(3)
    PK2650.HomeMotors()
    PK2650.ComputeOilIndex()
    PK2650.Measure()
PK2650.ComputeSlice()
Common.Pause()