# SERVER DEPLOYMENT GUIDE
# PLZ CV Engine - Centralized Monitoring
========================================================

## Overview

The application can now run in **SERVER MODE** - a single instance on a server monitoring a shared network location for all users.

---

## Architecture Options

### Option 1: Server-Based Monitoring (RECOMMENDED)
```
                    ┌─────────────────┐
                    │  SERVER PC      │
                    │  PLZ CV Engine  │
                    │  (Running 24/7) │
                    └────────┬────────┘
                             │
                    ┌────────▼────────┐
                    │ Network Share   │
                    │ \\SERVER\Data   │
                    └────────┬────────┘
                             │
        ┌────────────────────┼────────────────────┐
        │                    │                    │
   ┌────▼────┐          ┌────▼────┐         ┌────▼────┐
   │ User PC1│          │ User PC2│         │ User PC3│
   │(PowerAU)│          │(PowerAU)│         │(PowerAU)│
   └─────────┘          └─────────┘         └─────────┘
   
   Power Automate → Drop JSON → Server Monitors → Generate PDF → Email All
```

**Benefits:**
- ✓ Single point of maintenance
- ✓ Centralized logging
- ✓ No application on user PCs
- ✓ One email configuration
- ✓ Consistent processing

### Option 2: User-Based Monitoring (Original)
```
   Each user runs their own instance on their PC
   (Current OneDrive sync model)
```

---

## SERVER DEPLOYMENT STEPS

### Step 1: Set Up Shared Network Location

Create shared folder structure on server or NAS:

```
\\SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers\
├── Json\                    # Input folder (Power Automate writes here)
├── Templates\               # Excel template (place here manually)
├── Populated Template\      # Temporary working files
└── PDF CVs\                 # Output PDFs (users access here)
```

**Permissions:**
- Json folder: Write access for all users (Power Automate service account)
- Templates folder: Read-only for all, write for admin
- PDF CVs folder: Read access for all users
- Application service account: Full control on all folders

---

### Step 2: Configure Server Path

Edit `config.ini` on the server:

```ini
[PATHS]
# Use network UNC path for server deployment
root_path = \\SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers

# OR use local path if running on the same server as the share
root_path = D:\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers
```

**Examples:**

```ini
# Windows Server with shared drive
root_path = \\FILESERVER01\CompanyData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers

# Server local path (if app runs on file server itself)
root_path = E:\Shares\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers

# For testing on local PC (default if left empty)
root_path = 
```

---

### Step 3: Install on Server

1. **Copy deployment package to server:**
   ```
   C:\Program Files\PLZ CV Engine\
   ```

2. **Edit config.ini:**
   - Set `root_path` to network/local path
   - Configure email settings
   - Set `enabled = true`

3. **Install Excel on Server:**
   - Must have Microsoft Excel installed
   - Activate Excel license
   - Test Excel can open/save files

4. **Test manually first:**
   ```
   C:\Program Files\PLZ CV Engine\PLZ_CV_Engine.exe
   ```
   - Watch console for errors
   - Verify folder access
   - Check activity_log.txt

---

### Step 4: Configure as Windows Service (24/7 Operation)

#### Option A: Task Scheduler (Simple)

1. Open Task Scheduler on server
2. Create Task (not Basic Task):
   - Name: PLZ CV Engine
   - User: Service account with network access
   - ✓ Run whether user is logged on or not
   - ✓ Run with highest privileges
   - ✓ Configure for Windows Server 2016/2019/2022

3. Triggers:
   - At startup
   - At log on (optional)

4. Actions:
   - Start program: `C:\Program Files\PLZ CV Engine\PLZ_CV_Engine.exe`
   - Start in: `C:\Program Files\PLZ CV Engine\`

5. Conditions:
   - ✗ Start only if on AC power (uncheck)
   - ✗ Stop if on battery (uncheck)

6. Settings:
   - ✓ Allow task to be run on demand
   - If running task does not end when requested: Do not stop
   - Restart on failure: Every 5 minutes, up to 3 times

#### Option B: NSSM (Advanced - True Windows Service)

1. Download NSSM (Non-Sucking Service Manager): https://nssm.cc/

2. Install as service:
   ```powershell
   nssm install "PLZ CV Engine" "C:\Program Files\PLZ CV Engine\PLZ_CV_Engine.exe"
   nssm set "PLZ CV Engine" AppDirectory "C:\Program Files\PLZ CV Engine"
   nssm set "PLZ CV Engine" DisplayName "PLZ Collection Voucher Engine"
   nssm set "PLZ CV Engine" Description "Monitors and processes collection voucher PDFs"
   nssm set "PLZ CV Engine" Start SERVICE_AUTO_START
   nssm start "PLZ CV Engine"
   ```

3. Verify:
   ```powershell
   Get-Service "PLZ CV Engine"
   ```

---

### Step 5: Update Power Automate Flows

Update all Power Automate flows to write JSON files to the new network location:

**Old path (user OneDrive):**
```
C:\Users\[Username]\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers\Json\
```

**New path (network share):**
```
\\SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers\Json\
```

**In Power Automate:**
- Update "Create file" actions
- Point to network UNC path
- Ensure service account has write permissions

---

### Step 6: User Access to PDFs

Users access generated PDFs via:

1. **Network Share:**
   ```
   \\SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers\PDF CVs\
   ```

2. **Email:**
   - All configured recipients receive emails automatically
   - PDF attached to email

3. **Mapped Drive (Optional):**
   - Map network share to drive letter (e.g., P:)
   - Users browse P:\PDF CVs\

---

## MONITORING & MAINTENANCE

### Check Service Status

**Task Scheduler:**
```powershell
Get-ScheduledTask -TaskName "PLZ CV Engine"
```

**NSSM Service:**
```powershell
Get-Service "PLZ CV Engine"
nssm status "PLZ CV Engine"
```

### View Logs

- Activity log: `C:\Program Files\PLZ CV Engine\activity_log.txt`
- Windows Event Viewer (if using NSSM)
- Monitor file growth and rotate when large

### Restart Service

**Task Scheduler:**
```powershell
Stop-ScheduledTask -TaskName "PLZ CV Engine"
Start-ScheduledTask -TaskName "PLZ CV Engine"
```

**NSSM:**
```powershell
Restart-Service "PLZ CV Engine"
# OR
nssm restart "PLZ CV Engine"
```

---

## TESTING CHECKLIST

- [ ] Excel installed and licensed on server
- [ ] Network path accessible from server
- [ ] Service account has full permissions
- [ ] Template file exists in Templates folder
- [ ] Email configuration tested
- [ ] Test JSON placed in Json folder
- [ ] PDF generated in PDF CVs folder
- [ ] Email received by all recipients
- [ ] Service starts automatically on server reboot
- [ ] Activity log shows successful operations

---

## TROUBLESHOOTING

**Service won't start:**
- Check service account permissions
- Verify network path is accessible
- Review Windows Event Viewer
- Check activity_log.txt

**PDFs not generated:**
- Verify Excel installation on server
- Check Excel can run in background (no dialogs)
- Ensure template file exists and is accessible
- Review file permissions

**Network path issues:**
- Test UNC path from server: `Test-Path "\\SERVER\SharedData\..."`
- Verify DNS resolution
- Check firewall rules
- Test with service account credentials

**Email not sending:**
- Check SMTP port not blocked by firewall
- Verify email credentials in config.ini
- Test from server command line: `Test-NetConnection smtp.gmail.com -Port 587`

---

## SECURITY CONSIDERATIONS

1. **Service Account:**
   - Create dedicated service account
   - Grant minimum required permissions
   - Use strong password
   - Don't use admin account

2. **Network Security:**
   - Restrict Json folder to authorized users only
   - PDF folder read-only for general users
   - Monitor access logs

3. **Email Security:**
   - Use app passwords, not main password
   - Store config.ini securely (not in user-accessible location)
   - Regularly rotate credentials

4. **File Security:**
   - Enable auditing on shared folders
   - Monitor for unauthorized access
   - Regular backup of configurations

---

## BACKUP & RECOVERY

**Configuration Backup:**
```powershell
Copy-Item "C:\Program Files\PLZ CV Engine\config.ini" "\\BACKUP\Configs\PLZ_config_$(Get-Date -Format 'yyyyMMdd').ini"
```

**Service Recreation:**
1. Keep deployment package backed up
2. Document service account credentials
3. Export scheduled task: 
   ```powershell
   Export-ScheduledTask -TaskName "PLZ CV Engine" | Out-File "PLZ_Task.xml"
   ```

---

## SCALABILITY

Current design handles:
- ✓ Multiple users dropping files simultaneously
- ✓ File locking prevents duplicate processing
- ✓ Automatic retry on sync issues
- ✓ Recovery from crashes (.processing files)

**Performance:**
- Single Excel instance, sequential processing
- ~10-15 seconds per PDF
- Can handle 200-300 PDFs per hour

**If higher volume needed:**
- Deploy multiple server instances monitoring different folders
- Use queue-based architecture
- Consider Python-based PDF generation (no Excel dependency)

---

## SUMMARY

**Server Mode Benefits:**
✓ Centralized operation
✓ No user PC dependencies  
✓ Single point of configuration
✓ Easier maintenance
✓ Consistent logging
✓ Better control

**Deployment Checklist:**
1. [ ] Set up network share
2. [ ] Install Excel on server
3. [ ] Deploy application
4. [ ] Configure paths in config.ini
5. [ ] Set up Windows service
6. [ ] Update Power Automate flows
7. [ ] Test end-to-end
8. [ ] Monitor for 24 hours
9. [ ] Document for team

Ready for production! 🚀
