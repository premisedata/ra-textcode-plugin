# Office 365 SSO Setup Guide

## Complete These Steps to Enable Authentication

### 1. Register App in Azure Portal

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in:
   - **Name:** RA Doc Tagging Add-In (or your preferred name)
   - **Supported account types:** Accounts in this organizational directory only (Single tenant)
   - Click **Register**

### 2. Copy Application ID

1. On the app overview page, copy the **Application (client) ID**
2. Open `manifest.xml` in this repo
3. Replace `YOUR_APPLICATION_ID_HERE` (appears twice) with your actual Application ID:
   ```xml
   <Id>YOUR_APPLICATION_ID_HERE</Id>
   <Resource>api://premisedata.github.io/YOUR_APPLICATION_ID_HERE</Resource>
   ```

### 3. Configure API Permissions

1. In your Azure app, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**
3. Add these permissions:
   - `User.Read`
   - `profile`
   - `openid`
4. Click **Grant admin consent for [Your Org]**

### 4. Configure Authentication

1. Go to **Authentication** > **Add a platform**
2. Choose **Single-page application**
3. Add Redirect URI:
   ```
   https://premisedata.github.io/ra-textcode-plugin/taskpane.html
   ```
4. Under **Implicit grant and hybrid flows**, check:
   - ✅ Access tokens (used for implicit flows)
   - ✅ ID tokens (used for implicit and hybrid flows)
5. Click **Save**

### 5. Configure Expose an API (Optional but Recommended)

1. Go to **Expose an API**
2. Click **Set** next to Application ID URI
3. Accept the default: `api://YOUR_APPLICATION_ID`
4. Or customize to: `api://premisedata.github.io/YOUR_APPLICATION_ID`

### 6. Restrict Access to Authorized Users

Control who can use the add-in through Entra ID (Azure AD):

1. Go to **Entra ID** > **Enterprise Applications** > Find "RA Doc Tagging Add-In"
2. Click **Properties**
3. Set **"Assignment required?"** to **Yes**
4. Click **Save**
5. Go to **Users and groups** (in the left menu)
6. Click **Add user/group**
7. Select authorized users or security groups
8. Click **Assign**

**Result:** Only assigned users will be able to authenticate and use the add-in. Others will see an error.

### 7. Deploy Updated Files

1. Update `manifest.xml` with your Application ID (replace both instances of `YOUR_APPLICATION_ID_HERE`)
2. Commit and push to `gh-pages` branch
3. Wait for GitHub Actions to deploy (~2 minutes)
4. Update manifest in Word's wef folder:
   ```bash
   cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
   ```
5. Restart Word

### 8. Test Authentication

1. Open Word and load the RA Doc Tagging add-in
2. You should see "Authenticating..." briefly
3. You may be prompted to consent (first time only)
4. After success, you'll see "Signed in as: [Your Name]"
5. All coded text will now include your username

## Troubleshooting

**"Please sign in to Office 365"**
- Make sure you're signed into Office with your org account
- Try signing out and back in

**"Please grant consent to use this add-in"**
- Click consent button when prompted
- Or have admin pre-consent in Azure AD

**"Single Sign-On is not supported"**
- SSO requires Office 2016 or later
- Verify you're using a supported version

**User can't access add-in**
- Make sure they're assigned in Entra ID (Enterprise Applications > Users and groups)
- Verify "Assignment required?" is set to Yes
- Check API permissions are granted

**Still seeing "anonymous" username**
- Check browser console for errors (F12 in Word on Windows)
- Verify Application ID is correct in manifest
- Ensure API permissions are granted in Azure

## What Gets Tracked

When a user codes text, the system now stores:
- Username (email address)
- Timestamp
- Category
- All participant/demographic data
- Document filename
- Project information

This data appears in:
- Word comments on highlighted text
- Exported JSON files
- Future backend storage (Supabase/etc)
