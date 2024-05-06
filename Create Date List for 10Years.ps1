# Parameters
$SiteURL = "https://yourtenant.sharepoint.com/sites/yoursite"
$ListName = "YourListName"
$YearsToAdd = 10

# Connect to the SharePoint site
Connect-PnPOnline -Url $SiteURL -Interactive

# Create the list if it does not exist
$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    # Create the list
    $list = New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch
    # Add necessary columns
    Add-PnPField -List $ListName -DisplayName "Date" -InternalName "Date" -Type DateTime -AddToDefaultView
    Add-PnPField -List $ListName -DisplayName "Day of the Week" -InternalName "DayOfWeek" -Type Text -AddToDefaultView
    Add-PnPField -List $ListName -DisplayName "Is Work Day" -InternalName "IsWorkDay" -Type Choice -Choices "Yes","No" -AddToDefaultView
    Add-PnPField -List $ListName -DisplayName "Day Name" -InternalName "DayName" -Type Text -AddToDefaultView  # New column for the day name
}

# Date handling
$StartDate = Get-Date
$EndDate = $StartDate.AddYears($YearsToAdd)

# Loop through each day from today to the end of the period
while ($StartDate -le $EndDate) {
    # Determine the day of the week
    $DayOfWeek = $StartDate.DayOfWeek
    $DayName = $DayOfWeek.ToString()  # Convert DayOfWeek enum to string

    # Determine if it is a work day (Monday to Friday)
    $IsWorkDay = 'No'
    if ($DayOfWeek -ne 'Saturday' -and $DayOfWeek -ne 'Sunday') {
        $IsWorkDay = 'Yes'
    }

    # Add an item to the SharePoint list
    Add-PnPListItem -List $ListName -Values @{
        "Title" = $StartDate.ToString("MMMM d, yyyy")  # Formatted like 'May 5, 2024'
        "Date" = $StartDate
        "DayOfWeek" = $DayOfWeek
        "DayName" = $DayName  # Storing the day name
        "IsWorkDay" = $IsWorkDay
    }

    # Increment the date by one day
    $StartDate = $StartDate.AddDays(1)
}

# Disconnect the session
Disconnect-PnPOnline
