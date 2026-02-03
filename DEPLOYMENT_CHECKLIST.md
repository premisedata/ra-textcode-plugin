# CodeText Deployment Checklist

Quick reference for deploying CodeText to production.

---

## Pre-Deployment

- [ ] Test add-in locally and verify all functionality works
- [ ] Verify `test-categories.json` has all required categories
- [ ] Update version number in `taskpane.html`
- [ ] Update version in manifest files
- [ ] Review all documentation for accuracy

---

## AWS Setup

### S3 Bucket
- [ ] Create S3 bucket: `codetext-addin` (or your name)
- [ ] Enable static website hosting
- [ ] Configure CORS (see DEPLOYMENT_S3.md)
- [ ] Set bucket policy for public read
- [ ] Note the S3 website endpoint

### Upload Files
- [ ] Upload `taskpane-simple.js` (content-type: application/javascript)
- [ ] Upload `taskpane.html` (content-type: text/html)
- [ ] Upload `taskpane.css` (content-type: text/css)
- [ ] Upload `test-categories.json` (content-type: application/json)
- [ ] Upload `commands.html` (content-type: text/html)
- [ ] Make all files public

### CloudFront (Required for HTTPS)
- [ ] Create CloudFront distribution
- [ ] Set S3 website endpoint as origin
- [ ] Enable redirect HTTP to HTTPS
- [ ] Wait for deployment (5-15 minutes)
- [ ] Note CloudFront domain: `_______________.cloudfront.net`
- [ ] Test in browser: `https://YOUR-DOMAIN.cloudfront.net/taskpane.html`

---

## Manifest Configuration

- [ ] Copy `manifest-http.xml` to `manifest-production.xml`
- [ ] Replace ALL `localhost:3000` with CloudFront domain
- [ ] Update icon URLs (or use placeholders)
- [ ] Update provider name and support URL
- [ ] Verify all URLs use HTTPS
- [ ] Test manifest loads in browser
- [ ] Upload `manifest-production.xml` as `manifest.xml` to S3

**Manifest URL:** `https://YOUR-DOMAIN.cloudfront.net/manifest.xml`

---

## Testing

### Browser Test
- [ ] Open `https://YOUR-DOMAIN.cloudfront.net/taskpane.html`
- [ ] Verify page loads without errors
- [ ] Check browser console for errors
- [ ] Verify categories load

### Word Test (Windows)
- [ ] Download manifest.xml
- [ ] Install via Insert → Add-ins → Upload My Add-in
- [ ] Verify panel opens
- [ ] Test category selection
- [ ] Test text coding
- [ ] Test highlighting with colors
- [ ] Test report generation
- [ ] Test export functionality

### Word Test (Mac)
- [ ] Download manifest.xml
- [ ] Install via Insert → Add-ins → Add from file
- [ ] Verify panel opens
- [ ] Test all functionality (same as Windows)

---

## Distribution

### For Small Teams (Manual Install)
- [ ] Send manifest.xml URL to users
- [ ] Send USER_INSTALL_GUIDE.md to users
- [ ] Provide support email/contact

### For Large Organizations (Centralized Deployment)
- [ ] Access Microsoft 365 Admin Center
- [ ] Go to Settings → Integrated apps
- [ ] Upload custom app (manifest.xml)
- [ ] Assign to specific users/groups
- [ ] Monitor deployment status

**Recommended:** Use centralized deployment for 10+ users

---

## Documentation

- [ ] Update README.md with deployment date
- [ ] Document CloudFront domain for internal reference
- [ ] Create internal wiki page (if applicable)
- [ ] Set up support ticketing system
- [ ] Train support staff on common issues

---

## Post-Deployment

### Day 1
- [ ] Monitor for user installation issues
- [ ] Check CloudFront logs for errors
- [ ] Respond to user questions
- [ ] Document any issues

### Week 1
- [ ] Gather user feedback
- [ ] Monitor usage patterns
- [ ] Address any bugs
- [ ] Update documentation based on feedback

### Ongoing
- [ ] Regular backups of category files
- [ ] Monitor AWS costs
- [ ] Plan for updates/new features
- [ ] Review security settings quarterly

---

## Update Procedure

When you need to update the add-in:

1. **Update Local Files**
   - [ ] Make code changes
   - [ ] Test locally
   - [ ] Increment version number

2. **Upload to S3**
   - [ ] Upload modified files
   - [ ] Invalidate CloudFront cache

3. **Update Manifest (if needed)**
   - [ ] Update version in manifest
   - [ ] Upload new manifest
   - [ ] Notify users to reinstall (if manifest changed)

**Quick update command:**
```bash
aws s3 sync /path/to/TaskPaneMod s3://codetext-addin/ \
  --exclude "*.md" --exclude "*.py"
aws cloudfront create-invalidation --distribution-id YOUR-ID --paths "/*"
```

---

## Rollback Procedure

If something goes wrong:

- [ ] Revert files in S3 to previous version (use S3 versioning)
- [ ] Invalidate CloudFront cache
- [ ] Test functionality
- [ ] Notify users if needed

**Enable S3 versioning** for easy rollback!

---

## Support Information

### User Support Contacts
- Primary: ____________________
- Secondary: __________________
- Email: ______________________

### Technical Contacts
- AWS Admin: __________________
- IT Support: __________________
- Developer: ___________________

### Important URLs
- Manifest: `https://__________.cloudfront.net/manifest.xml`
- Add-in URL: `https://__________.cloudfront.net/taskpane.html`
- S3 Bucket: `s3://codetext-addin/`
- CloudFront ID: `______________`

---

## Common Issues & Solutions

### Issue: Add-in won't install
- Check manifest URL is accessible
- Verify HTTPS (not HTTP)
- Test in different browser
- Check Word version (2016+)

### Issue: Categories not loading
- Check `test-categories.json` is uploaded
- Verify file is public
- Check content-type is `application/json`
- Clear CloudFront cache

### Issue: Updates not appearing
- Invalidate CloudFront cache
- Clear Word cache
- Close and reopen Word
- Check version number updated

---

## Security Checklist

- [ ] HTTPS enforced (CloudFront redirect)
- [ ] CORS properly configured
- [ ] S3 bucket not writable by public
- [ ] No sensitive data in add-in files
- [ ] Manifest from trusted source only
- [ ] Regular security reviews scheduled

---

## Cost Monitoring

Estimated monthly AWS costs:
- S3 Storage: $0.01 - $1 (depending on file size)
- CloudFront: $0.10 - $5 (depending on traffic)
- Data Transfer: $0.10 - $2

**Total: ~$0.50 - $10/month** for typical usage

- [ ] Set up AWS billing alerts
- [ ] Monitor CloudFront usage
- [ ] Review costs monthly

---

## Success Criteria

Deployment is successful when:
- [ ] 90%+ of users successfully install
- [ ] No critical bugs reported
- [ ] All core features working
- [ ] Reports generating correctly
- [ ] User satisfaction >80%
- [ ] Support tickets <5 per week

---

## Version History

| Version | Date | Changes | Deployed By |
|---------|------|---------|-------------|
| 1.33    | 2025-01-25 | Initial production release | ___________ |
|         |      |         |             |
|         |      |         |             |

---

**Deployment Lead:** ___________________
**Deployment Date:** ___________________
**Next Review Date:** ___________________

