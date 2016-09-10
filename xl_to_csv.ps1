# Initialize a couple of array variables.
$sam_account_names = @()
$array_to_export = @()

# Codes for savin printers to associate users with the correct
# initials.
$SAVIN_TITLE1_CODES = @{
    "AD" = "1";
    "CD" = "2";
    "EF" = "3";
    "GH" = "4";
    "IJK" = "5";
    "LMN" = "6";
    "OPQ" = "7";
    "RST" = "8";
    "UVW" = "9";
    "XYZ" = "10";
}

# Helper function that obtains users via AD.
function get-ad-user{
    $sam_acct_v1 = $args[0]
    $sam_acct_v2 = $args[1]

    $user = ""
    $error_message = ""

    Try {
        $user = Get-ADUser -Identity $sam_acct_v1 -Properties EmailAddress
    }
    Catch {
        $error_message = $_.Exception.Message
        $error_message >> dot_notation_sam_error.txt
        "User " + $sam_acct_v1 + " was not found..." >> dot_notation_sam_error.txt
    }

    #$user
    #$user -eq $null

    if ($user.length -eq 0){
        Try {
            $user = Get-ADUser -Identity $sam_acct_v2 -Properties EmailAddress
        }
        Catch {
            $error_message = $_.Excepation.Message
            $error_message >> compact_notation_sam_error.txt
            "User " + $sam_acct_v2 + " was not found..." >> compact_notation_sam_error.txt
        }
    }

    return $user
}

# Helper function which creates an array of the relevant user info to import 
# to the savin printer.
function set-user-info-array{
    $user_info_array = @()
    $user_info = ""
    $key_display = ""
    $title1_num_value = ""

    $sam_acct_v1 = $args[0]
    $sam_acct_v2 = $args[1]

    #$user_info = Get-ADUser -Identity $sam_acct -Properties EmailAddress
    $user_info = get-ad-user $sam_acct_v1 $sam_acct_v2
    Try {
        $key_display = $user_info.GivenName[0] + " " + $user_info.Surname
    }
    Catch {
        $error_message = $_.Exception.Message
        $sam_acct_v1 >> null_error.txt
        $sam_acct_v2 >> null_error.txt
        $error_message >> null_error.txt
    }
        
    foreach ($key in $SAVIN_TITLE1_CODES.KEYS.GetEnumerator()){
        Try {
            if ($key.Contains($user_info.Surname[0])){
                    $title1_num_value = $SAVIN_TITLE1_CODES[$key]
            }
        }
        Catch {
            $error_message = $_.Exception.Message
            $error_message >> null_error.txt
        }
    }

    $user_info_array = $user_info.Name, $key_display, $user_info.EmailAddress, $title1_num_value
    return $user_info_array
}

# Import csv by prompting user to enter the path to the csv file.
$csv_sheet = Import-Csv $(Read-Host "Enter path to csv file")

# Loop that goes through the name column in the csv file.
foreach ($name in $csv_sheet.Name){
    $sam_account_name = ""
    $split_name = ""

    if ($name -match '\ \(\w*\)'){
        $name = $name -replace '\ \(\w*\)', ""
    }
    elseif ($name -match '\(\w*\)'){
        $name = $name -replace '\(\w*\)', ""
    }

    if ($name.Contains("*")){
        $split_name = $name.split(", ")
        $sam_account_name_v1 = $split_name[2].substring(0,$split_name[2].length - 1) + "." + $split_name[0]
        $sam_account_name_v2 = $split_name[2][0] + $split_name[0]
        
        $array_to_export += set-user-info-array $sam_account_name_v1 $sam_account_name_v2
    }
    else{
        $split_name = $name.Split(", ")
        $sam_account_name_v1 = $split_name[2] + "." + $split_name[0]
        $sam_account_name_v2 = $split_name[2][0] + $split_name[0]

        $array_to_export += set-user-info-array $sam_account_name_v1 $sam_account_name_v2
    }
}

$array_to_export
