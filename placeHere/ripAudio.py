import os
import shutil
import zipfile

audioExtensions = (".aiff", ".au", ".mid", ".midi", ".mp3", ".m4a", ".mp4", ".wav", ".wma")
startDir = os.getcwd()
fileList = os.listdir()
audioOutputDir = os.path.join(startDir, "audios")
os.makedirs(audioOutputDir, exist_ok=True)

for file in fileList:
    if not (file.endswith(".zip") or file.endswith(".pptx")):
        continue
    baseName, ext = os.path.splitext(file)
    originalName = baseName
    zipFileName = baseName + ".zip"
    if ext != ".zip":
        os.rename(file, zipFileName)
    else:
        zipFileName = file
    with zipfile.ZipFile(zipFileName, "r") as myZip:
        if myZip.testzip() is not None:
            print(f"ERROR in {zipFileName}")
            continue
        for member in myZip.namelist():
            if member.endswith(audioExtensions):
                myZip.extract(member)
    mediaPath = os.path.join(startDir, "ppt", "media")
    if not os.path.exists(mediaPath):
        print(f"NO AUDIO in {zipFileName}")
        continue
    for audioFile in os.listdir(mediaPath):
        srcPath = os.path.join(mediaPath, audioFile)
        destPath = os.path.join(audioOutputDir, audioFile)
        if os.path.exists(destPath):
            os.remove(destPath)
        os.rename(srcPath, destPath)
    os.chdir(startDir)
    shutil.rmtree(os.path.join(startDir, "ppt"), ignore_errors=True)
    if ext != ".zip":
        os.rename(zipFileName, originalName + ".pptx")
