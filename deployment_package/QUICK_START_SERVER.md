# QUICK START: SERVER DEPLOYMENT
========================================================

## 🎯 You are here: Converting to SERVER MODE

**What changed:**
- ❌ OLD: Each PC runs the application, monitors local OneDrive
- ✅ NEW: ONE server runs the application, monitors shared network location

---

## 🚀 3-STEP SERVER SETUP

### **Step 1: Setup Network Share (5 minutes)**

1. Create shared folder on your server:
   ```
   \\YOUR-SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers\
   ```

2. Set permissions:
   - Everyone: Read
   - Service Account: Full Control
   - Power Automate: Write (Json folder)

3. Copy the template file to:
   ```
   \\YOUR-SERVER\SharedData\...\Templates\Collection Voucher Template.xlsx
   ```

---

### **Step 2: Configure & Install (5 minutes)**

**Option A: Automated Setup (Recommended)**
```powershell
cd deployment_package
.\setup_server.ps1
```
Follow the prompts to configure paths and email.

**Option B: Manual Setup**
1. Edit `config.ini`:
   ```ini
   [PATHS]
   root_path = \\YOUR-SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers
   
   [EMAIL]
   enabled = true
   sender_email = plzpopservices@gmail.com
   sender_password = wxtg rjnh mpin ujwl
   recipients = ejojo@premiumzimbabwe.com, tnyakurukwa@premiumzimbabwe.com, vmukandatsama@premiumzimbabwe.com
   ```

2. Copy deployment_package to server:
   ```
   C:\Program Files\PLZ CV Engine\
   ```

3. Test: Double-click `PLZ_CV_Engine.exe`

---

### **Step 3: Make it Run 24/7 (5 minutes)**

**Quick method:** Task Scheduler
```powershell
# Run as Administrator
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PLZ CV Engine\PLZ_CV_Engine.exe" -WorkingDirectory "C:\Program Files\PLZ CV Engine"
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "PLZ CV Engine" -Action $action -Trigger $trigger -Principal $principal
```

**Done!** Application runs on server boot.

---

## 📋 UPDATE POWER AUTOMATE

Update your Power Automate flows to write JSON to the network path:

**Change from:**
```
C:\Users\[Username]\Premium Leaf Zimbabwe\...\Json\
```

**Change to:**
```
\\YOUR-SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers\Json\
```

---

## ✅ TESTING

1. **Drop a test JSON** in the Json folder
2. **Watch for PDF** in PDF CVs folder (should appear in ~15 seconds)
3. **Check email** - all 3 recipients should receive it
4. **Verify log** at `C:\Program Files\PLZ CV Engine\activity_log.txt`

---

## 🎯 PATH EXAMPLES

**Your specific setup (customize):**
```ini
# If app runs on file server itself:
root_path = D:\Shares\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers

# If app runs on different server accessing network:
root_path = \\FILESERVER\PLZData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers

# For testing locally (uses OneDrive):
root_path = 
```

---

## 📊 CURRENT EMAIL CONFIGURATION

**From:** plzpopservices@gmail.com

**To:**
- ejojo@premiumzimbabwe.com
- tnyakurukwa@premiumzimbabwe.com  
- vmukandatsama@premiumzimbabwe.com

**Subject:** Collection Voucher {REQ-NUMBER} Generated

**Body:**
```
Hello,

A new Collection Voucher PDF has been generated.

Voucher Number: {REQ-NUMBER}
Generated: {TIMESTAMP}

Please find the PDF attached.

Best regards,
Savannah Stores
```

---

## 🔧 TROUBLESHOOTING

**Error: Cannot access path**
- Verify network path from server: `Test-Path "\\SERVER\..."`
- Check service account has permissions
- Try accessing path in File Explorer

**Error: Excel not found**
- Install Microsoft Excel on the server
- Activate Excel license
- Test: Open Excel manually on server

**PDFs not generating**
- Check template file exists in Templates folder
- Verify Excel is licensed and activated
- Review `activity_log.txt` for errors

**Email not working**
- Verify SMTP port 587 not blocked
- Check credentials in config.ini
- Test: `Test-NetConnection smtp.gmail.com -Port 587`

---

## 📚 FULL DOCUMENTATION

For complete details, see:
- `SERVER_DEPLOYMENT.md` - Full server setup guide
- `DEPLOYMENT_INSTRUCTIONS.txt` - General deployment
- `README.txt` - Quick reference

---

## 💡 ARCHITECTURE COMPARISON

### OLD (User Mode):
```
User PC 1 → Local OneDrive → PLZ App → PDF → Email
User PC 2 → Local OneDrive → PLZ App → PDF → Email  
User PC 3 → Local OneDrive → PLZ App → PDF → Email
```

### NEW (Server Mode):
```
User PC 1 ──┐
User PC 2 ──┼─→ Network Share → SERVER (PLZ App) → PDF → Email (All)
User PC 3 ──┘
```

**Benefits:**
✓ Single app instance
✓ One configuration
✓ Centralized logging
✓ No app on user PCs
✓ Easier maintenance

---

**Ready? Run `setup_server.ps1` to begin!** 🚀
