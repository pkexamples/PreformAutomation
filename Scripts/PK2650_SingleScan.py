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

PK2650.SetPreformId('Test', 18.25)
PK2650.MoveZManual('Load the preform')
PK2650.LocateTip()
PK2650.MoveToZ(50)
PK2650.EquilibratePreform(5)
PK2650.NewSlice()
Focus(3)
PK2650.HomeMotors()
PK2650.ComputeOilIndex()
PK2650.Measure()
PK2650.ComputeSlice()
Common.Pause()