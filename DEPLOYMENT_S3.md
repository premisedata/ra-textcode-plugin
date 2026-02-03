# CodeText Word Add-in - S3 Deployment Guide

Complete guide for deploying the CodeText Word Add-in to AWS S3 for production use.

---

## Part 1: Prepare Files for Deployment

### 1.1 Files to Upload

The following files need to be uploaded to S3:

```
TaskPaneMod/
├── taskpane-simple.js      ← Main JavaScript file
├── taskpane.html            ← Main HTML file
├── taskpane.css             ← Stylesheet
├── test-categories.json     ← Category data
├── commands.html            ← Empty file for manifest
└── manifest-production.xml  ← Production manifest (created in step 2)
```

### 1.2 Create Production Manifest

Create a production version of the manifest with your S3 URL:

```bash
cd /Users/jason/c0d3t3xt/TaskPaneMod
cp manifest-http.xml manifest-production.xml
```

---

## Part 2: Set Up S3 Bucket

### 2.1 Create S3 Bucket

1. Log in to AWS Console
2. Go to S3 service
3. Click **Create bucket**
4. Bucket settings:
   - **Bucket name**: `codetext-addin` (or your preferred name)
   - **Region**: Choose closest to your users (e.g., `us-east-1`)
   - **Block Public Access**: **UNCHECK** "Block all public access"
   - ⚠️ Acknowledge that the bucket will be public
5. Click **Create bucket**

### 2.2 Enable Static Website Hosting

1. Click on your bucket name
2. Go to **Properties** tab
3. Scroll to **Static website hosting**
4. Click **Edit**
5. Settings:
   - **Static website hosting**: Enable
   - **Hosting type**: Host a static website
   - **Index document**: `taskpane.html`
6. Click **Save changes**
7. **Note the endpoint URL** (e.g., `http://codetext-addin.s3-website-us-east-1.amazonaws.com`)

### 2.3 Configure CORS

Word Add-ins require CORS to be configured.

1. Go to **Permissions** tab
2. Scroll to **Cross-origin resource sharing (CORS)**
3. Click **Edit**
4. Paste this CORS configuration:

```json
[
    {
        "AllowedHeaders": [
            "*"
        ],
        "AllowedMethods": [
            "GET",
            "HEAD"
        ],
        "AllowedOrigins": [
            "*"
        ],
        "ExposeHeaders": [],
        "MaxAgeSeconds": 3000
    }
]
```

5. Click **Save changes**

### 2.4 Set Bucket Policy

Make the bucket publicly readable:

1. Go to **Permissions** tab
2. Scroll to **Bucket policy**
3. Click **Edit**
4. Paste this policy (replace `codetext-addin` with your bucket name):

```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Sid": "PublicReadGetObject",
            "Effect": "Allow",
            "Principal": "*",
            "Action": "s3:GetObject",
            "Resource": "arn:aws:s3:::codetext-addin/*"
        }
    ]
}
```

5. Click **Save changes**

---

## Part 3: Upload Files to S3

### 3.1 Upload via AWS Console

1. Click on your bucket
2. Click **Upload**
3. Click **Add files**
4. Select these files:
   - `taskpane-simple.js`
   - `taskpane.html`
   - `taskpane.css`
   - `test-categories.json`
   - `commands.html`
5. Click **Upload**

### 3.2 Set Content Types (Important!)

S3 needs to know the correct MIME types:

1. For **each file**, click on it
2. Click **Actions** → **Edit metadata**
3. Set the following:

| File | Content-Type |
|------|-------------|
| `taskpane-simple.js` | `application/javascript` |
| `taskpane.html` | `text/html` |
| `taskpane.css` | `text/css` |
| `test-categories.json` | `application/json` |
| `commands.html` | `text/html` |

4. Click **Save changes** for each

### 3.3 Upload via AWS CLI (Alternative)

If you have AWS CLI installed:

```bash
cd /Users/jason/c0d3t3xt/TaskPaneMod

# Upload files with correct content types
aws s3 cp taskpane-simple.js s3://codetext-addin/ --content-type "application/javascript"
aws s3 cp taskpane.html s3://codetext-addin/ --content-type "text/html"
aws s3 cp taskpane.css s3://codetext-addin/ --content-type "text/css"
aws s3 cp test-categories.json s3://codetext-addin/ --content-type "application/json"
aws s3 cp commands.html s3://codetext-addin/ --content-type "text/html"

# Make files public
aws s3api put-object-acl --bucket codetext-addin --key taskpane-simple.js --acl public-read
aws s3api put-object-acl --bucket codetext-addin --key taskpane.html --acl public-read
aws s3api put-object-acl --bucket codetext-addin --key taskpane.css --acl public-read
aws s3api put-object-acl --bucket codetext-addin --key test-categories.json --acl public-read
aws s3api put-object-acl --bucket codetext-addin --key commands.html --acl public-read
```

---

## Part 4: Configure CloudFront (HTTPS) - REQUIRED

⚠️ **Word Add-ins REQUIRE HTTPS.** S3 website endpoints only provide HTTP, so we need CloudFront.

### 4.1 Create CloudFront Distribution

1. Go to **CloudFront** in AWS Console
2. Click **Create distribution**
3. Origin settings:
   - **Origin domain**: Your S3 website endpoint (e.g., `codetext-addin.s3-website-us-east-1.amazonaws.com`)
   - **Protocol**: HTTP only (S3 website endpoints don't support HTTPS)
   - **Name**: Auto-filled
4. Default cache behavior:
   - **Viewer protocol policy**: Redirect HTTP to HTTPS
   - **Allowed HTTP methods**: GET, HEAD
   - **Cache policy**: CachingOptimized
5. Settings:
   - **Price class**: Use all edge locations (or choose based on budget)
   - **Alternate domain name (CNAME)**: Leave empty (or add custom domain if you have one)
6. Click **Create distribution**
7. **Wait 5-15 minutes** for deployment
8. **Note your CloudFront domain** (e.g., `d111111abcdef8.cloudfront.net`)

### 4.2 Test CloudFront

Once deployed, test in browser:

```
https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/taskpane.html
```

You should see the CodeText interface load.

---

## Part 5: Update Manifest with Production URLs

### 5.1 Edit manifest-production.xml

Open `/Users/jason/c0d3t3xt/TaskPaneMod/manifest-production.xml` and replace all localhost URLs with your CloudFront domain:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
           xsi:type="TaskPaneApp">
  <Id>12345678-1234-1234-1234-123456789abc</Id>
  <Version>1.33.0.0</Version>
  <ProviderName>Your Organization</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="CodeText"/>
  <Description DefaultValue="Qualitative data coding tool for Microsoft Word"/>
  
  <!-- REPLACE WITH YOUR CLOUDFRONT DOMAIN -->
  <IconUrl DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/assets/icon-64.png"/>
  
  <SupportUrl DefaultValue="https://your-support-site.com"/>
  
  <AppDomains>
    <AppDomain>https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net</AppDomain>
  </AppDomains>
  
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/taskpane.html"/>
  </DefaultSettings>
  
  <Permissions>ReadWriteDocument</Permissions>
  
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
          
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="CommandsGroup.Label" DefaultValue="CodeText"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show CodeText"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open CodeText coding panel"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

### 5.2 Upload Manifest to S3

```bash
aws s3 cp manifest-production.xml s3://codetext-addin/manifest.xml --content-type "application/xml" --acl public-read
```

Or via console, upload as `manifest.xml`.

---

## Part 6: Create Add-in Icons (Optional but Recommended)

Create simple icon images:

1. Create three PNG files:
   - `icon-16.png` (16x16 pixels)
   - `icon-32.png` (32x32 pixels)
   - `icon-80.png` (80x80 pixels)

2. Create an `assets` folder in S3:
   - In S3, create a folder called `assets`
   - Upload the three icon files to this folder
   - Set content-type to `image/png` for each
   - Make them public

Or use placeholder URL in manifest if you don't have icons yet.

---

## Part 7: Distribute to Users

### 7.1 Download Manifest URL

Users need the manifest URL. This is:

```
https://YOUR-CLOUDFRONT-DOMAIN.cloudfront.net/manifest.xml
```

### 7.2 Distribution Methods

**Option A: Email Instructions**

Send users this link and installation instructions (see Part 8).

**Option B: Internal SharePoint/Portal**

Host the manifest.xml file on your internal portal and provide download link.

**Option C: Centralized Deployment (Microsoft 365 Admin)**

If you have Microsoft 365 admin access:

1. Go to Microsoft 365 admin center
2. Navigate to **Settings** → **Integrated apps**
3. Click **Upload custom apps**
4. Upload your `manifest-production.xml`
5. Assign to users/groups
6. Users will see the add-in automatically in Word

This is the best method for enterprise deployment - no manual installation needed!

---

## Part 8: User Installation Instructions

Create this guide for your users:

### Installing CodeText Add-in for Microsoft Word

#### For Windows Users:

1. Open Microsoft Word
2. Go to **Insert** tab → **Add-ins** → **Get Add-ins**
3. Click **My Add-ins** on the left
4. At the bottom, click **Upload My Add-in**
5. Click **Browse** and select the `manifest.xml` file you downloaded
6. Click **Upload**
7. The CodeText panel should appear on the right side

#### For Mac Users:

1. Open Microsoft Word
2. Go to **Insert** tab → **Add-ins**
3. Click **My Add-ins**
4. At the bottom, click **+ Add a custom add-in** → **Add from file**
5. Select the `manifest.xml` file you downloaded
6. Click **Add**
7. The CodeText panel should appear on the right side

#### Accessing CodeText:

Once installed, you can open CodeText from:
- **Home** tab → **CodeText** group → **Show CodeText**

Or refresh the panel:
- **Insert** → **Add-ins** → **My Add-ins** → **CodeText**

---

## Part 9: Updating the Add-in

When you make changes to the code:

### 9.1 Update Files

1. Edit your local files in `TaskPaneMod/`
2. Increment version number in:
   - `taskpane.html` (update version display and cache-busting `?v=XX`)
   - `manifest-production.xml` (update `<Version>` tag)

### 9.2 Upload to S3

```bash
# Upload updated files
aws s3 cp taskpane-simple.js s3://codetext-addin/ --content-type "application/javascript"
aws s3 cp taskpane.html s3://codetext-addin/ --content-type "text/html"

# Invalidate CloudFront cache (important!)
aws cloudfront create-invalidation --distribution-id YOUR-DIST-ID --paths "/*"
```

### 9.3 Users Get Updates

- Users will get updates automatically when they reload Word
- CloudFront caching might delay updates (use invalidation to force)
- If manifest changes, users need to reinstall (or use centralized deployment)

---

## Part 10: Troubleshooting

### Issue: Add-in won't load

**Check:**
- Is CloudFront distribution deployed?
- Can you access `https://YOUR-DOMAIN.cloudfront.net/taskpane.html` in browser?
- Are CORS headers configured in S3?
- Are files public (check bucket policy)?

### Issue: "This app cannot be run" error

**Solution:**
- Manifest must use HTTPS URLs
- Test manifest URL is accessible in browser
- Check Word's add-in error logs

### Issue: Changes not appearing

**Solution:**
- Clear CloudFront cache
- Clear Word's cache (close and reopen Word)
- Check version numbers are updated

### Issue: Categories not loading

**Solution:**
- Verify `test-categories.json` is uploaded and public
- Check browser console for 404 errors
- Verify JSON file has correct content-type

---

## Part 11: Security Considerations

### 11.1 Data Privacy

- All data is stored locally in the user's browser (no server-side storage)
- Reports are generated locally in the user's Word document
- No data is transmitted to external servers

### 11.2 Access Control

- The S3 bucket is publicly readable (required for Word Add-ins)
- Consider using CloudFront signed URLs for additional security
- Use Microsoft 365 centralized deployment to control who has access

### 11.3 HTTPS

- Always use HTTPS (CloudFront handles this)
- Never use HTTP for production Word Add-ins

---

## Summary Checklist

- [ ] S3 bucket created and configured
- [ ] Static website hosting enabled
- [ ] CORS configured
- [ ] Bucket policy set for public read
- [ ] Files uploaded with correct content-types
- [ ] CloudFront distribution created and deployed
- [ ] Manifest updated with CloudFront URLs
- [ ] Manifest tested (loads in browser)
- [ ] Icons created and uploaded (optional)
- [ ] User installation instructions prepared
- [ ] Test installation on Windows and Mac

---

## Quick Reference

**Your CloudFront URL:** `https://YOUR-DIST-ID.cloudfront.net`

**Manifest URL for users:** `https://YOUR-DIST-ID.cloudfront.net/manifest.xml`

**Test the add-in:** Open manifest URL in browser - should download XML file

**Update command:**
```bash
aws s3 sync /Users/jason/c0d3t3xt/TaskPaneMod s3://codetext-addin/ \
  --exclude "*.md" --exclude "*.py" --exclude ".git/*" \
  --exclude "manifest-http.xml" --exclude "manifest.xml"
aws cloudfront create-invalidation --distribution-id YOUR-DIST-ID --paths "/*"
```

---

**Need help?** Check AWS documentation or contact your IT administrator.

