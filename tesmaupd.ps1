# TESMA Asset auto-update Script

# TESMA Online https://tesmaonline.chg-meridian.com/
# TESMA API https://api.chg-meridian.com/docs/index
# CHG-MERIDIAN https://www.chg-meridian.com/

# Created by Fernando Almeida (04/07/2019) - fernando.almeida@chg-meridian.com

$VERSION = "1.0.1"
$SHOW_PUT_RESULT = $false # Change to $true to show PUT result

# Customer info
$CUSTOMER_CostCenter = "BRADESCO1"
$CUSTOMER_Local = "Predio Branco"
$CUSTOMER_Department = "Contabilidade"

# TESMA
$TESMA_CustomerID = "333435"
$TESMA_AssetID = ""

$TESMA_Username = "01333435c" # Webservice enabled user 
$TESMA_Password = "nfx32ng47" # Password

$TESMA_AuthorizationKey = ""
$TESMA_API_Uri = "https://tesmademo.chg-meridian.com/api/assets"
$TESMA_Header_Accept = "application/vnd.tesma.v1+json"
$TESMA_Content_Type = "application/vnd.tesma.v1+json"

$ConfigFile = ($env:USERPROFILE + "\TESMA.XML") # File is located on user profile. If deleted, system will query TESMA API again and recreate XML file.

function SaveAssetID {
    # Save asset ID to XML file
    # SaveAssetID(<assetid>)

    param (
        $assetid
    )

    # Create new XML object
    $xmlsettings = New-Object System.Xml.XmlWriterSettings
    $xmlsettings.Indent = $true
    $xmlsettings.IndentChars = "    "

    $XmlWriter = [System.XML.XmlWriter]::Create($ConfigFile, $xmlsettings)
    
    $xmlWriter.WriteStartDocument()
    
    $xmlWriter.WriteStartElement("Config")
    $xmlWriter.WriteElementString("AssetID", $assetid)
    $xmlWriter.WriteEndElement()
    
    $xmlWriter.WriteEndDocument()
    $xmlWriter.Flush()
    $xmlWriter.Close()
}
function GetManufacturer {
    # Return the computer manufacturer from CIM (https://docs.microsoft.com/en-us/windows/win32/wmisdk/common-information-model)
    # $return = GetManufacturer()

    return (Get-CimInstance -ClassName Win32_BIOS).Manufacturer
}
function GetComputerModel {
    # Return the computer model from CIM (https://docs.microsoft.com/en-us/windows/win32/wmisdk/common-information-model)
    # $return = GetComputerModel()

    return (Get-CimInstance -ClassName Win32_ComputerSystem).Model
}
function GetSerialNumber {
    # Return the computer serial number from CIM (https://docs.microsoft.com/en-us/windows/win32/wmisdk/common-information-model)
    # $return = GetSerialNumber()

    return (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
}
function GetUserName {
    # Return the current user from environment (https://docs.microsoft.com/en-us/dotnet/api/system.environment.username?view=netframework-4.8)
    # $return = GetUserName()

    return [Environment]::UserName
}
function GetUserDomain {
    # Return the current user domain from environment (https://docs.microsoft.com/en-us/dotnet/api/system.environment.username?view=netframework-4.8)
    # $return = GetUserDomain()

    return [Environment]::UserDomainName
}
function GetComputerName {
    # Return the computer name (or hostname) from environment (https://docs.microsoft.com/en-us/dotnet/api/system.environment.username?view=netframework-4.8)
    # $return = GetComputerName()

    return [Environment]::MachineName
}
function GetOSVersion {
    # Return the operating system version from WMI (https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page)
    # $return = GetOSVersion()

    return (Get-WmiObject -class Win32_OperatingSystem).Caption
}
function GetOSBits {
    # Return the system type bits from environment (https://docs.microsoft.com/en-us/dotnet/api/system.environment.username?view=netframework-4.8)
    # $return = GetOSBits()

    if([Environment]::Is64BitOperatingSystem) {
        return "64 Bit"
    } else {
        return "32 Bit"
    }
}
function GetComputerType {
    # Return the computer type from WMI (https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page)
    # $return = GetComputerType()

    $tmp = (Get-WMIObject -Class Win32_ComputerSystem -Property PCSystemType).PCSystemType
    switch($tmp)
    {
        1 { return "Desktop" }
        2 { return "Notebook" }
        3 { return "Workstation" }
        4 { return "Servidor" }
        5 { return "Servidor (SOHO)" }
        default { return "Desconhecido" }
    }
}
function AssetIdBySerialNumber {
    # Return the asset ID from TESMA, by given computer serial number
    # $return = AssetIdBySerialNumber()

    param (
        $serial
    )

    if([System.IO.File]::Exists($ConfigFile)) { # Does TESMA.XML exist? If yes, use asset ID from XML, in order to save 1 TESMA API query
        # Open TESMA.XML
        $xml = [xml](Get-Content $ConfigFile)
        # Read asset ID from XML file
        $tmp = $xml.Config.AssetID
        Write-Host ("ID read from XML: " + $tmp)
    } else { # No TESMA.XML - maybe this is the first time using the script
        try {
            # Query TESMA API
            $tmp = (Invoke-RestMethod -Uri ("$TESMA_API_Uri`?`$filter=serial_number eq '$serial'") -Headers @{Authorization=("Basic {0}" -f $TESMA_AuthorizationKey);Accept=$TESMA_Header_Accept}).Id
            Write-Host ("ID read from TESMA: " + $tmp)
            # Save asset ID to new TESMA.XML local file
            SaveAssetID $tmp
        }
        catch { # Catch error. Show details and finish with error (-1)
            Write-Host ("Error: " + $_)
            exit -1
        }
    }

    return $tmp
}
function GetAssetInformationById {
    # Return the asset information from TESMA, by given asset ID
    # $return = AssetIdBySerialNumber(<assetid>)

    param (
        $assetid
    )

    return Invoke-RestMethod -Uri ("$TESMA_API_Uri/$TESMA_CustomerID/$assetid") -Headers @{Authorization=("Basic {0}" -f $TESMA_AuthorizationKey);Accept=$TESMA_Header_Accept}
}
function GetAssetInformationBySerialNumber {
    # Return the asset information from TESMA, by given computer serial number
    # $return = AssetIdBySerialNumber(<serialnumber>)

    param (
        $serial
    )

    return Invoke-RestMethod -Uri ("$TESMA_API_Uri?`$filter=serial_number eq '$serial'") -Headers @{Authorization=("Basic {0}" -f $TESMA_AuthorizationKey);Accept=$TESMA_Header_Accept}
}

# Welcome message
Write-Output ":: TESMA Asset info auto-update Script ($VERSION)`nCreated by Fernando Almeida (04/07/2019) - fernando.almeida@chg-meridian.com`n"

# Encode Base64 username and password
$Bytes = [System.Text.Encoding]::UTF8.GetBytes($TESMA_Username + ":" + $TESMA_Password)
$TESMA_AuthorizationKey = [Convert]::ToBase64String($Bytes) # Create TESMA Base64 Authorization key

# Read the asset ID by given serial number
$TESMA_AssetID = AssetIdBySerialNumber(GetSerialNumber)

if([String]::IsNullOrEmpty($TESMA_AssetID)) {
    Write-Output "Error: Asset not found in TESMA."
    exit -1
} else {
    # Create information package
    $Data = @{
        SerialNumber = (GetSerialNumber);
        Description = (GetComputerModel);
        Location = $CUSTOMER_Department;
        CostCentre = $CUSTOMER_CostCenter;
        OrderId = "";
        Factory = $CUSTOMER_Local;
        Reference = "";
        Text1 = ((GetUserDomain) + "\" + (GetUserName));
        Text2 = (GetComputerName);
        Text3 = (GetManufacturer);
        Text4 = (GetOSVersion);
        Text5 = (GetOSBits);
        Text6 = (GetComputerType);
        Text7 = "";
        Text8 = "";
        Text9 = "";
        Text10 = "";
        Number1 = 0;
        Number2 = 0;
        Number3 = 0;
        Number4 = 0;
        Number5 = 0;
        Date1 = (Get-Date -UFormat "%Y-%m-%dT%H:%M:%S");
        Date2 = "1901-01-01T00:00:00";
        Date3 = "1901-01-01T00:00:00";
        Date4 = "1901-01-01T00:00:00";
        Date5 = "1901-01-01T00:00:00";
    }

    # Convert to JSON format
    $json_data = $Data | ConvertTo-Json

    try {
        if($SHOW_PUT_RESULT) {
            (Invoke-RestMethod -Method PUT -Uri ("$TESMA_API_Uri/$TESMA_CustomerID/$TESMA_AssetID") -ContentType $TESMA_Content_Type -Headers @{Authorization=("Basic {0}" -f $TESMA_AuthorizationKey);Accept=$TESMA_Header_Accept;ContentType=$TESMA_Content_Type} -Body $json_data)
        } else {
            $result = (Invoke-RestMethod -Method PUT -Uri ("$TESMA_API_Uri/$TESMA_CustomerID/$TESMA_AssetID") -ContentType $TESMA_Content_Type -Headers @{Authorization=("Basic {0}" -f $TESMA_AuthorizationKey);Accept=$TESMA_Header_Accept;ContentType=$TESMA_Content_Type} -Body $json_data)
        }
    }
    catch { # Catch error
        $errormessage = ($_ | ConvertFrom-Json)

        #Parse error messagem
        if($errormessage.Message -eq "The request is invalid.") { # Request invalid. Show error message and finish with error (-1)
            Write-Output "Error: Unable to update TESMA. Check CustomerID and AssetID`n"
            exit -1
        }

        if($errormessage.Message -eq "Missing credentials") { # Missing credentials. Show error message and finish with error (-1)
            Write-Output "Error: Unable to update TESMA. Missing credentials`n"
            exit -1
        }

        if($errormessage.Message -eq "Invalid credentials") { # Invalid credentials. Show error message and finish with error (-1)
            Write-Output "Error: Unable to update TESMA. Invalid credentials`n"
            exit -1
        }
    }
    
    # All good. Show message and finish with success (0)
    Write-Output "TESMA updated sucessfully.`n"
    exit 0
}
