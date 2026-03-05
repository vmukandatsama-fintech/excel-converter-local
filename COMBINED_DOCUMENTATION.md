# Combined Documentation (Local Deployment)

This consolidated guide is for local deployment only.

---

## Source: DEPLOYMENT_GUIDE.md

# PLZ CV ENGINE - DEPLOYMENT GUIDE

## ✅ Deployment Package Ready!

The deployment package has been created at:
**`deployment_package/`**

---

## 📦 Package Contents

```
deployment_package/
├── PLZ_CV_Engine.exe              # Standalone executable
├── config.ini                     # Configuration file
├── Start_PLZ_CV_Engine.bat        # Quick start script
└── README.txt                     # Quick reference
```

---

## 🚀 Quick Deployment Steps

### **Step 1: Distribute Package**
1. Copy the **entire `deployment_package/` folder** to each target PC
2. Recommended location: `C:\Program Files\PLZ CV Engine\`
3. Or use a shared file location for easy distribution

### **Step 2: Configure Email (Each PC)**
Edit `config.ini`:
```ini
[EMAIL]
enabled = true
smtp_server = smtp.gmail.com
smtp_port = 587
sender_email = plzpopservices@gmail.com
sender_password = <app-password>
recipients = vmukandatsama@gmail.com, vmukandatsama@premiumzimbabwe.com
```

### **Step 3: Launch Application**
- Double-click **`PLZ_CV_Engine.exe`** OR
- Double-click **`Start_PLZ_CV_Engine.bat`**
- Application runs in system tray

### **Step 4: Verify**
- Check system tray for PLZ icon
- Verify folder created: `C:\Users\[Username]\Premium Leaf Zimbabwe\...`
- Test by placing JSON file in `Json/` folder

---

## 🔄 Auto-Start Setup (Optional)

### Method A: Task Scheduler (Recommended)
```
1. Win+R → taskschd.msc
2. Create Basic Task → Name: "PLZ CV Engine"
3. Trigger: At log on
4. Action: Start program → Browse to PLZ_CV_Engine.exe
5. Run with highest privileges
```

### Method B: Startup Folder
```
1. Win+R → shell:startup
2. Create shortcut to PLZ_CV_Engine.exe
3. Paste shortcut in opened folder
```

---

## 📝 Testing Checklist

- [ ] Application starts without errors
- [ ] System tray icon visible
- [ ] JSON detection works
- [ ] PDF generation successful
- [ ] Email sending works (if enabled)
- [ ] Activity log created

---

## 🛠️ Maintenance

### Update Deployment
1. Rebuild: `python build_deployment.py`
2. Replace `PLZ_CV_Engine.exe` in `deployment_package`
3. Keep existing `config.ini`

### View Logs
- Check `activity_log.txt` next to `.exe`
- Monitor for errors

---

## 📞 Support

**Application won't start:**
- Run as Administrator
- Check Windows Defender exclusions
- Verify Excel is installed

**PDF not generated:**
- Check template exists
- Verify Excel installation
- Review `activity_log.txt`

**Email not sending:**
- Verify credentials in `config.ini`
- Check firewall/antivirus SMTP rules

---

**Deployment Package Location:**
`C:\Users\VictorMukandatsama\Development\ExcelConverter\deployment_package\`
