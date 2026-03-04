# PowerShell-Graph-CreateABookingRoomEvent.ps1
# This script demonstrates how to create a calendar event (meeting) in another user's calendar (e.g. a room mailbox) using Microsoft Graph API.
# Generated initially with CoPilot.
# Prompt: create a powershell graph sample to create a calendar event in another users calendar, such as a room mailbox.  Add verbose comments.
<#
.SYNOPSIS
  Creates a calendar event that requests/Books a room by adding the room mailbox as a resource attendee.

.DESCRIPTION
  This script:
    1) Acquires an app-only token (client credentials).
    2) Creates an event in the organizer's mailbox via:
         POST /v1.0/users/{organizerUPN}/events
    3) Adds the room mailbox as an attendee with attendee type "resource".
       If the room mailbox is configured to AutoAccept, it will accept/decline automatically.

.NOTES
  Permissions (Application):
    - Calendars.ReadWrite

  Behavior depends on room mailbox settings:
    - AutoAccept on => room may accept automatically
    - Conflict rules, booking windows, etc., can cause declines

  Tip:
    - To see acceptance status later, read the event and check attendees' status.
#>

# ----------------------------
# 1) CONFIGURATION (EDIT THESE)
# ----------------------------

# Azure AD tenant ID (GUID)
$TenantId     = "00000000-0000-0000-0000-000000000000"

# App (client) ID
$ClientId     = "11111111-1111-1111-1111-111111111111"

# Client secret (store securely in production)
$ClientSecret = "YOUR_CLIENT_SECRET_VALUE"

# Organizer mailbox where the event will be created
# (Room booking is requested by inviting the room as a resource attendee)
$OrganizerUpn = "organizer@contoso.com"

# Room mailbox SMTP address (resource mailbox)
$RoomSmtp     = "confroom101@contoso.com"

# Event details
$Subject      = "Room booking test via Graph"
$BodyText     = "Created by PowerShell/Graph. This event requests the room resource."

# Time window (use ISO 8601). You can use local offset or Z for UTC.
# Example: Eastern time with -05:00 offset (adjust as needed).
$StartDateTime = "2026-03-05T10:00:00-05:00"
$EndDateTime   = "2026-03-05T10:30:00-05:00"

# Time zone preference (affects how Graph returns times; payload below uses explicit offsets)
$PreferredTimeZone = "Eastern Standard Time"

# ----------------------------
# 2) GET APP-ONLY ACCESS TOKEN
# ----------------------------

Write-Verbose "Requesting app-only access token (client credentials)..." -Verbose

$TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

$TokenBody = @{
  client_id     = $ClientId
  client_secret = $ClientSecret
  grant_type    = "client_credentials"
  scope         = "https://graph.microsoft.com/.default"
}

try {
  $TokenResponse = Invoke-RestMethod -Method POST -Uri $TokenEndpoint -Body $TokenBody -ContentType "application/x-www-form-urlencoded"
}
catch {
  throw "Failed to get token. $($_.Exception.Message)"
}

$AccessToken = $TokenResponse.access_token
if (-not $AccessToken) { throw "Token response did not contain access_token." }

Write-Verbose "Access token acquired." -Verbose

# ----------------------------
# 3) BUILD EVENT PAYLOAD
# ----------------------------

# IMPORTANT:
# - Room is added as an attendee with type "resource"
# - location is set for readability (does not itself book the room; the attendee does)
# - showAs / responseRequested can be set as desired
# - isReminderOn / reminderMinutesBeforeStart optional

$EventPayload = @{
  subject = $Subject
  body    = @{
    contentType = "text"
    content     = $BodyText
  }

  start = @{
    dateTime = $StartDateTime
    timeZone = $PreferredTimeZone
  }

  end = @{
    dateTime = $EndDateTime
    timeZone = $PreferredTimeZone
  }

  location = @{
    displayName = "Conference Room 101"
  }

  attendees = @(
    @{
      emailAddress = @{
        address = $RoomSmtp
        name    = "Conference Room 101"
      }
      type = "resource"   # Key part for room/resource attendee
    }
  )

  # Request responses (room may auto-process anyway depending on config)
  responseRequested = $true

  # Optional: mark as busy in organizer calendar
  showAs = "busy"
}

# Convert to JSON. Use enough depth to include nested objects.
$EventJson = $EventPayload | ConvertTo-Json -Depth 10

Write-Verbose "Event payload JSON:" -Verbose
Write-Verbose $EventJson -Verbose

# ----------------------------
# 4) CREATE EVENT IN ORGANIZER'S MAILBOX
# ----------------------------

# Endpoint: POST /users/{organizer}/events
$CreateEventUrl = "https://graph.microsoft.com/v1.0/users/$OrganizerUpn/events"

$Headers = @{
  Authorization = "Bearer $AccessToken"
  Accept        = "application/json"
  "Content-Type"= "application/json"
  Prefer        = "outlook.timezone=`"$PreferredTimeZone`""
}

Write-Verbose "Creating event in organizer mailbox: $OrganizerUpn" -Verbose
Write-Verbose "POST $CreateEventUrl" -Verbose

try {
  $CreatedEvent = Invoke-RestMethod -Method POST -Uri $CreateEventUrl -Headers $Headers -Body $EventJson
}
catch {
  # If room booking fails due to policy, the event may still be created (and room may decline later).
  throw "Create event call failed. $($_.Exception.Message)"
}

# ----------------------------
# 5) OUTPUT RESULTS
# ----------------------------

Write-Verbose "Event created successfully." -Verbose
Write-Output ("Created Event Id: {0}" -f $CreatedEvent.id)
Write-Output ("Subject        : {0}" -f $CreatedEvent.subject)

# Optional: show attendee status snapshot (may update after room processes the request)
if ($CreatedEvent.attendees) {
  Write-Output ""
  Write-Output "Attendee status (initial):"
  $CreatedEvent.attendees | ForEach-Object {
    $addr   = $_.emailAddress.address
    $type   = $_.type
    $status = $_.status.response
    Write-Output (" - {0} [{1}] => {2}" -f $addr, $type, $status)
  }
}

Write-Output ""
Write-Output "NOTE: The room acceptance/decline may occur asynchronously depending on room mailbox settings."
