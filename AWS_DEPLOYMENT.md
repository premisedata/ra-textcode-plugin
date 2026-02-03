# AWS Deployment Guide for CodeText Word Add-in

This guide covers deploying the Word Add-in to AWS with HTTPS.

## Prerequisites

- AWS account
- Domain name (optional, but recommended)
- SSL certificate (AWS Certificate Manager can provide this)

## Option 1: S3 + CloudFront (Recommended for Static Hosting)

### Step 1: Create S3 Bucket

1. Go to S3 Console
2. Create a new bucket (e.g., `codetext-addin`)
3. **Disable** "Block all public access" (or configure bucket policy for public read)
4. Enable "Static website hosting"
5. Set index document to `taskpane.html`

### Step 2: Upload Files

Upload all files from `TaskPaneMod` folder to the S3 bucket:
- `manifest.xml`
- `taskpane.html`
- `taskpane.js`
- `taskpane.css`
- `python/functions.py` (optional, for reference)

### Step 3: Create CloudFront Distribution

1. Go to CloudFront Console
2. Create distribution
3. Set origin to your S3 bucket
4. Configure:
   - **Viewer Protocol Policy**: Redirect HTTP to HTTPS
   - **Allowed HTTP Methods**: GET, HEAD, OPTIONS
   - **Cache Policy**: CachingDisabled (or custom for development)
5. Add your domain (if you have one)
6. Request SSL certificate in AWS Certificate Manager if needed

### Step 4: Update manifest.xml

Replace all `https://localhost:3000` URLs with your CloudFront domain:

```xml
<SourceLocation DefaultValue="https://your-cloudfront-domain.cloudfront.net/taskpane.html"/>
<AppDomain>https://your-cloudfront-domain.cloudfront.net</AppDomain>
```

### Step 5: Configure CORS (if needed)

If you encounter CORS errors, add this bucket policy:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "PublicReadGetObject",
      "Effect": "Allow",
      "Principal": "*",
      "Action": "s3:GetObject",
      "Resource": "arn:aws:s3:::your-bucket-name/*"
    }
  ]
}
```

## Option 2: AWS Amplify (Simpler Setup)

### Step 1: Connect Repository

1. Go to AWS Amplify Console
2. Click "New app" > "Host web app"
3. Connect your Git repository
4. Point to `TaskPaneMod` folder

### Step 2: Configure Build

Amplify will auto-detect static files. No build step needed.

### Step 3: Update manifest.xml

Update URLs to your Amplify domain:
```xml
<SourceLocation DefaultValue="https://your-app-id.amplifyapp.com/taskpane.html"/>
```

## Option 3: EC2 with Nginx/Apache

### Step 1: Launch EC2 Instance

1. Launch EC2 instance (Ubuntu recommended)
2. Configure security group to allow HTTPS (port 443)

### Step 2: Install Web Server

**Nginx:**
```bash
sudo apt update
sudo apt install nginx
sudo apt install certbot python3-certbot-nginx
```

**Apache:**
```bash
sudo apt update
sudo apt install apache2
sudo apt install certbot python3-certbot-apache
```

### Step 3: Upload Files

Upload `TaskPaneMod` folder contents to `/var/www/html/` (or your web root)

### Step 4: Configure SSL

```bash
sudo certbot --nginx  # or --apache
```

### Step 5: Update manifest.xml

Update URLs to your EC2 domain/IP

## Testing Deployment

1. **Verify HTTPS**: All URLs must use `https://` (not `http://`)
2. **Test manifest.xml**: Upload to Word and check for errors
3. **Check browser console**: Open Word, press F12, check for CORS or loading errors
4. **Verify CDN resources**: Ensure Pyodide and DuckDB load from CDN

## Common Issues

### CORS Errors

Add CORS headers to your web server:
```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET, POST, OPTIONS
Access-Control-Allow-Headers: Content-Type
```

### Manifest.xml Not Loading

- Ensure manifest.xml is accessible via HTTPS
- Check that all URLs in manifest.xml use HTTPS
- Verify AppDomain matches your domain

### Pyodide/DuckDB Not Loading

- These load from CDN, so ensure internet connectivity
- Check browser console for blocked resources
- Consider hosting WASM files locally if CDN is blocked

## Production Considerations

1. **Custom Domain**: Use your own domain instead of CloudFront/Amplify default
2. **Caching**: Configure CloudFront cache for better performance
3. **Monitoring**: Set up CloudWatch alarms for errors
4. **Backup**: Keep manifest.xml and source files in version control
5. **Updates**: Update manifest.xml version when deploying changes

## Cost Estimate

- **S3 + CloudFront**: ~$1-5/month (depending on traffic)
- **Amplify**: Free tier available, then pay-as-you-go
- **EC2**: ~$10-20/month for t2.micro instance

## Security Notes

- Word Add-ins require HTTPS (enforced by Office.js)
- Keep manifest.xml URLs updated when domain changes
- Consider using AWS WAF for additional protection
- Regularly update dependencies (Pyodide, DuckDB versions)

