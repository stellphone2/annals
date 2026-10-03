[app]

# (str) Title of your application
title = Annals

# (str) Package name - only letters, digits, underscores
package.name = annals

# (str) Package domain (reverse-DNS style)
package.domain = org.smo

# (str) Source code where the main.py lives
source.dir = .

# (list) Source files to include (let buildozer pick everything)
source.include_exts = py,png,jpg,kv,atlas,json,ttf

# (list) Patterns to exclude from the APK
source.exclude_dirs = tests, bin, .git, __pycache__, .buildozer

# (str) App version
version = 1.0

# -----------------------------------------------------------------------
# Requirements
# All Python packages your app imports must be listed here.
# 'android' is the pyjnius/python-for-android glue; kivy includes it.
# -----------------------------------------------------------------------
requirements = python3,kivy==2.3.0,openpyxl,python-docx,fpdf2,androidstorage4kivy

# (str) The icon of the application
# icon.filename = %(source.dir)s/icon.png

# (str) Supported orientation: landscape, portrait, sensorLandscape, sensorPortrait, all
orientation = portrait

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0

# (int) Target Android API
android.api = 33

# (int) Minimum Android API your APK will run on (Android 5.0+, covers ~99% of devices)
android.minapi = 21

# (str) Android build-tools version to use — pin to avoid licence issues with latest
android.build_tools_version = 34.0.0

# (int) Android NDK version to use (full revision string)
android.ndk = 25b

# (bool) Automatically accept Android SDK licences
android.accept_sdk_license = True

# (str) Android NDK directory (optional; buildozer downloads it automatically)
# android.ndk_path =

# (str) Android SDK directory (optional; buildozer downloads it automatically)
# android.sdk_path =

# (list) Permissions your app needs
android.permissions = WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, INTERNET

# (bool) Enable AndroidX (required by modern Kivy builds)
android.enable_androidx = True

# (str) Gradle build backend  (gradle is the modern default)
android.gradle_dependencies =

# (str) Python-for-android branch / version
# p4a.branch = master

# (list) Android additional libraries to copy into libs/
# android.add_libs_armeabi_v7a = libs/android/*.so

# (str) The entry-point class for python-for-android
# (leave blank for the default Kivy bootstrap)
# android.entrypoint = org.kivy.android.PythonActivity

# -----------------------------------------------------------------------
# Build settings
# -----------------------------------------------------------------------

# (str) Log level (0 = error, 1 = info, 2 = debug)
log_level = 2

# (int) Display warning on missing wheels
warn_on_root = 1

# (str) Output directory for the .apk / .aab
# buildozer_dir = .buildozer

[buildozer]

# (int) Log level (0 = error, 1 = info, 2 = debug)
log_level = 2

# (int) Display warning on missing wheels
warn_on_root = 1
