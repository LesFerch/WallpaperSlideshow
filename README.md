# WallpaperSlideshow

### Version 1.0.1

Multi-monitor, multi-folder wallpaper slideshows. Command line or GUI.

## How to Download and Install

**Note**: Some antivirus software may falsely detect the download as a virus. This can happen any time you download a new executable and may require extra steps to whitelist the file.

### Install Using Setup Program

[![image](https://github.com/user-attachments/assets/75e62417-c8ee-43b1-a8a8-a217ce130c91)Download the installer](https://github.com/LesFerch/WallpaperSlideshow/releases/download/1.0.1/WallpaperSlideshow-Setup.exe)

The WallpaperSlideshow installer requires administrator access.

1. Download the installer using the link above.
2. Double-click **WallpaperSlideshow-Setup.exe** to start the installation.
3. In the SmartScreen window, click **More info** and then **Run anyway**.

**Note**: The installer is only provided in English, but the program works with any language.

### Portable Use

[![image](https://github.com/LesFerch/WinSetView/assets/79026235/0188480f-ca53-45d5-b9ff-daafff32869e)Download the zip file](https://github.com/LesFerch/WallpaperSlideshow/releases/download/1.0.1/WallpaperSlideshow.zip)

Using WallpaperSlideshow as a portable app does NOT require administrator access.

1. Download the zip file using the link above.
2. Extract the contents.
3. Move the contents to a permanent location of your choice. For example **C:\Tools\WallpaperSlideshow**.
4. Double-click **WallpaperSlideshowUI.exe** to configure and start the slideshow.
5. In the SmartScreen window, click **More info** and then **Run anyway**.

## Summary

WallpaperSlideshow changes your wallpaper, like the built-in Windows slideshow, with the major difference being that a different image folder can be assigned to each monitor. This allows for images to be matched to the hardware (e.g. portrait vs landscape) or to have different sets of images shown on each display.

If you only have only monitor, you may still find WallpaperSlideshow useful for its folder watch feature.

Full feature list:
 
 - Image folders can be set individually for each monitor. 
 - The image wait time can be set individually for each monitor.
 - Wait times can range from 1 second to 100 hours.
 - Image folders are watched. New files get added to the slideshow. Deleted files are removed from the slideshow.
 - The slideshow start image is randomly selected from the files in the image folder.
 - Any monitor can have a static image by pointing it to a folder with a single image.
 - Any monitor's wallpaper can be managed manually, or by another program, by pointing it to a folder with no images.
 - Keeps running when monitors are connected or disconnected.
 - Survives an Explorer restart.

WallpaperSlideshow is meant to be simple and lightweight with support for command line operation and a simple GUI. The CPU and memory footprint is very small. It is not meant to compete with big programs like Wallpaper Engine.

**Note**: The built-in Windows slideshow does a nice crossfade when it changes images. Unfortunately, Microsoft provides no API for that feature, so WallpaperSlideshow only does an instant change of wallpaper images in order to remain lightweight.

## How to Use (GUI)

Light:

<img width="799" height="194" alt="image" src="https://github.com/user-attachments/assets/a1e967b6-9a50-44af-a9ce-51f98063313a" />



\
Dark:

<img width="801" height="196" alt="image" src="https://github.com/user-attachments/assets/6ed860d9-2cd4-4294-8d7c-0dab56289e34" />


### Interface

- The settings for each monitor are displayed in a grid with monitor numbers that match the Windows `Display settings`.

- Click the folder icon (📁) to select an image folder for each monitor. You can also click on the text and directly type or paste a folder path. Do not include quotes.

- To have a single image displayed on a monitor, set it to a folder that contains that one image.

- To exclude a monitor from the slideshow, and be able to set its wallpaper manually (or using another program), set it to a folder that contains no images. An empty folder is a good choice in this case.

- Click the stopwatch icon (⏱) to set a time in hours, minutes, seconds to wait between image changes. You can also click on the number and directly type in a value in seconds. The wait time is shown in seconds. The minimum is 1 and the maximum is 359999 (~ 100 hours).

- To have the slideshow run when you login, enable `Run at startup`.

- To close the GUI without making any changes,  click the `X`.

- To stop the slideshow (kill the WallpaperSlideshow.exe process), click `  Exit  `. 

- Click the `  OK  ` button to run (or continue) the slideshow with your new settings. If an invalid folder path was entered, the GUI will remain on screen and the invalid path will be highlighted.

**Note**: The monitor numbers are shown 1-based in the GUI, but are saved 0-based in the registry.

## How to Use (command line)

### wallpaperslideshow [folder1] [seconds1] [folder2] [seconds2] ...

folder1 = image folder for monitor 1\
seconds1 = wait time between images for monitor 1 in seconds\
and so on...

- Of course, unlike the GUI, folder paths that contain spaces must be quoted!

- If a folder-wait pair is not provided for a connected monitor, it will use the settings for the first monitor.

- Settings are written to `HKCU\Software\WallpaperSlideshow`

- If no arguments are provided, settings from the registry will be used.

### wallpaperslideshow /x

The `/x` option will kill the `wallpaperslideshow.exe` process.
 
\
\
[![image](https://github.com/LesFerch/WinSetView/assets/79026235/63b7acbc-36ef-4578-b96a-d0b7ea0cba3a)](https://github.com/LesFerch/WallpaperSlideshow)





