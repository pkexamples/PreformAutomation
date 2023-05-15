##Service Name=Index_Profile_ASCII
##Friendly Name=Auto-save ASCII Index Profile
"""
Extract_Index_Profile_to_ASCII

PKSL Persistor to export index profile to ASCII text file
"""

import clr, math, os
clr.AddReference('PK.Configuration', 'PK.PreformAnalysisBase', 'PK.InstrumentBase')
from PhotonKinetics.Common.Configuration import ConfigurationServer
from PhotonKinetics.Measurement.PreformAnalysis import PreformMeasurement, PreformSlice, PreformScan, RefractiveIndexProfile
from System.Globalization.CultureInfo import CurrentCulture, InvariantCulture


# Field separator
#sep = CurrentCulture.TextInfo.ListSeparator
#sep = InvariantCulture.TextInfo.ListSeparator
sep = '\t'

# Set file extension to match separator
#fileExt = 'csv'
#fileExt = 'csv'
fileExt = 'txt'

# Unit conversion
M_TO_MM = 1e3
M_TO_Microns = 1e6


def ExportIndexProfileToText(rip, dirName):
    # Build file name
    sample_id = rip.ID.GetElementByName('Preform ID')
    loc_part = GetZAndWString(rip)
    file_name_parts = (sample_id, rip.ProfileType, loc_part, fileExt)
    fileName = '{0}_{1}_{2}.{3}'.format(*file_name_parts)
    filePath = os.path.join(dirName, fileName)
    with open(filePath, 'w') as f:
        # Header
        f_zPos = rip.ZPosition * M_TO_MM
        zPos = '{0:6.2f}'.format(round(f_zPos))
        f_wPos = math.degrees(rip.WPosition)
        wPos = '{0:6.2f}'.format(round(f_wPos))
        f.write('{}\n'.format(rip.FormattedDescription(sep)))
        for i in rip.ID.IDCollection:
            f.write(sep.join((i.Prompt, i.Entry,'')))
        f.write(sep.join((rip.ID.Timestamp.ToString(), rip.ID.StartTime.ToString(), rip.ID.OperatorName, rip.SerialNumber)))
        f.write('\n')
        f.write(sep.join(('Radius (mm)', 'Refractive Index')))
        f.write('\n')
        # Write profile
        for p in rip.Profile:
            f.write('{1:5.3f}{0}{2:7.6f}\n'.format(sep, p.Radius * M_TO_MM, p.Index))

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
        # Create folder for text files
        directoryNameBase = os.path.join(
        ConfigurationServer.Path['Preform Analysis', 'Index Profile ASCII Files'],
        m.Result.ID.GetElementByName('Preform ID'))
        # Ensure unique name
        directoryName = GetUniqueDirName(directoryNameBase)
        # Create folder
        os.mkdir(directoryName)
        # Make folder name for average profiles
        avgProfileDirectoryNameBase = '{0}_{1}'.format(directoryName, 'AverageProfiles')
        avgProfileDirectoryName = GetUniqueDirName(avgProfileDirectoryNameBase)
        os.mkdir(avgProfileDirectoryName)
        # Export for each average profile for slices and individual scans within slices
        for profile in m.Result:
            try:
                ExportIndexProfileToText(profile, avgProfileDirectoryName)
            except:
                pass
            if isinstance(profile, PreformSlice):
                # Export each profile in the `PreformSlice`
                for scan in profile:
                    try:
                        ExportIndexProfileToText(scan, directoryName)
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
            directoryName = directoryNameBase + "_ScanIndexProfile_{}".format(loc_part)
            directoryName = GetUniqueDirName(directoryName)
            os.mkdir(directoryName)
            ExportIndexProfileToText(profile, directoryName)
        elif isinstance(profile, PreformSlice):
            loc_part = GetZString(profile)
            directoryName = directoryNameBase + "_SliceIndexProfile_{}".format(loc_part)
            directoryName = GetUniqueDirName(directoryName)
            os.mkdir(directoryName)
            for scan in profile:
                # Ensure any errors don't prevent subsequent scans from extracting data
                try:
                    ExportIndexProfileToText(scan, directoryName)
                except:
                    pass
