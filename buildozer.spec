[app]
title = Annals
package.name = annals
package.domain = org.annals
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json
version = 1.0

# Pure-Python only: no python-docx / lxml / fpdf2 / pillow needed any more
requirements = python3,kivy==2.3.0,pyjnius,android,androidstorage4kivy,openpyxl,et_xmlfile,certifi

orientation = portrait
fullscreen = 0

# Permissions: INTERNET for the Google sheet; storage only for the optional crash log
android.permissions = INTERNET, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE
android.api = 34
android.minapi = 24
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True
android.release_artifact = apk

# androidstorage4kivy (share sheet / shared storage) needs AndroidX
android.enable_androidx = True

[buildozer]
log_level = 2
warn_on_root = 1
