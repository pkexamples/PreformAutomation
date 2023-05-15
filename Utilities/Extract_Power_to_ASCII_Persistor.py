##Service Name=Power_ASCII
##Friendly Name=Save Power as ASCII
"""
Extract_Power_to_ASCII

PKSL Persistor to export scan power to ASCII text file
"""
import clr, math, os
clr.AddReference('PK.Configuration', 'PK.PreformAnalysisBase', 'PK.InstrumentBase')
from PhotonKinetics.Common.Configuration import ConfigurationServer
from System.Globalization.CultureInfo import CurrentCulture
from PhotonKinetics.Measurement.PreformAnalysis import PreformMeasurement, PreformSlice, PreformScan, RefractiveIndexProfile
from System.Globalization.CultureInfo import CurrentCulture, InvariantCulture


# string formatting
#sep = CurrentCulture.TextInfo.ListSeparator 
#sep = InvariantCulture.TextInfo.ListSeparator
sep = '\t'

# file extension
# change to match separator
#fileExt = 'csv'
fileExt = 'txt'

# unit conversion
M_TO_MM = 1e3
M_TO_Microns = 1e6

#Declare function to export single profile to text
def ExportPowerToText(rip, dirName):
    # build file name
    sample_id = rip.ID.GetElementByName('Preform ID')
    loc_part = GetZAndWString(rip)
    file_name_parts = (sample_id, 'Power', loc_part, fileExt)
    fileName = '{0}_{1}_{2}.{3}'.format(*file_name_parts)
    filePath = os.path.join(dirName, fileName)
    with open(filePath, 'w') as f:
        # Header
        f.write('PK2650 Power Data\n')
        f_zPos = rip.ZPosition * M_TO_MM
        zPos = '{0:.3f}'.format(round(f_zPos))
        f_wPos = math.degrees(rip.WPosition)
        wPos = '{0:.2f}'.format(round(f_wPos))
        f.write(sep.join(('Z Position (mm)', zPos, 'W Position (deg)', wPos + '\n')))
        for i in rip.ID.IDCollection:
            f.write(i.Prompt + sep)
        f.write(sep.join(('Time Stamp', 'Start Time', 'Operator', 'Serial Number\n')))
        for i in rip.ID.IDCollection:
            f.write(i.Entry + sep)
        f.write(sep.join((rip.ID.Timestamp.ToString(), rip.ID.StartTime.ToString(), rip.ID.OperatorName, rip.SerialNumber + '\n')))
        #Deflection
        f.write('{}\t{}\n'.format('Radius (mm)', 'Power'))
        if rip.MeasurementState.YScan:
            step_mm = rip.YScan.StepSize * 1e3
            pwr = rip.YScan.Power
            if rip.MeasurementState.DeflectionFunction:
                shiftCenter = rip.CenterIndex + rip.TruncationLeftEdge
                for j in range(rip.TruncationLeftEdge, rip.TruncationRightEdge - 1, 1):
                    f.write(sep.join(('{0:.3f}','{1:.1f}\n')).format(step_mm * (j - shiftCenter), pwr[j]))
            else:
                for j in range(len(pwr)):
                    f.write(sep.join(('{0:.3f}','{1:.1f}\n')).format(step_mm * (j - rip.CenterIndex), pwr[j]))
        else:
            f.write('No data present\n')

def GetZString(rip):
    return '{0:d}'.format(int(round(rip.ZPosition * M_TO_MM)))

def GetZAndWString(rip):
    return "{}_{}".format(GetZString(rip), int(round(math.degrees(rip.WPosition))))

def GetUniqueDirName(dir_name):
    dName = dir_name
    counter = 0
    while os.path.exists(dName):
        counter += 1
        dName = '{0}_{1}'.format(dir_name, counter)
    return dName

for m in MarkedObjects:
    if isinstance(m.Result, PreformMeasurement):
        # create folder for text files
        directoryNameBase = os.path.join(
            ConfigurationServer.Path['Preform Analysis', 'Index Profile ASCII Files'],
            m.Result.ID.GetElementByName('Preform ID') + '_Power')

        # ensure unique name
        directoryName = GetUniqueDirName(directoryNameBase)

        # create folder
        os.mkdir(directoryName)
        # export deflection function for each individual scan found
        for profile in m.Result:
            if isinstance(profile, PreformScan):
                #Ensure any errors don't prevent subsequent scans from extracting data
                try:
                    ExportPowerToText(profile, directoryName)
                except:
                    pass
            elif isinstance(profile, PreformSlice):
                for scan in profile:
                    #Ensure any errors don't prevent subsequent scans from extracting data
                    try:
                        ExportPowerToText(scan, directoryName)
                    except:
                        pass
    elif isinstance(m.Result, RefractiveIndexProfile):
        # create folder for text files
        directoryNameBase = os.path.join(
            ConfigurationServer.Path['Preform Analysis', 'Index Profile ASCII Files'],
            m.Result.ID.GetElementByName('Preform ID'))
        profile = m.Result
        if isinstance(profile, PreformScan):
            loc_part = GetZAndWString(profile)
            directoryName = directoryNameBase + "_ScanPower_{}".format(loc_part)
            directoryName = GetUniqueDirName(directoryName)
            os.mkdir(directoryName)
            ExportPowerToText(profile, directoryName)
        elif isinstance(profile, PreformSlice):
            loc_part = GetZString(profile)
            directoryName = directoryNameBase + "_SlicePower_{}".format(loc_part)
            directoryName = GetUniqueDirName(directoryName)
            os.mkdir(directoryName)
            for scan in profile:
                # Ensure any errors don't prevent subsequent scans from extracting data
                try:
                    ExportPowerToText(scan, directoryName)
                except:
                    pass
