(function () {
    app.beginUndoGroup("Organize Project");

    var proj = app.project;

    if (!proj) {
        alert("No project open!");
        return;
    }
    var config = {
        folders: {
            videos: "Videos",
            images: "Images",
            audio: "Audio",
            compositions: "Compositions",
            solids: "Solids",
            precomps: "Precomps",
            unusedComps: "Comps unused",
            footage: "Footage",
            other: "Other"
        },
        videoExts: [".mp4", ".mov", ".avi", ".mkv", ".mxf", ".r3d", ".m4v", ".mpg", ".mpeg"],
        imageExts: [".jpg", ".jpeg", ".png", ".tif", ".tiff", ".psd", ".ai", ".svg", ".gif", ".bmp", ".exr"],
        audioExts: [".mp3", ".wav", ".aac", ".m4a", ".aiff", ".flac", ".ogg"]
    };


    function getOrCreateFolder(name, parentFolder) {
        var parent = parentFolder || proj.rootFolder;

        // Chercher si le dossier existe déjà
        for (var i = 1; i <= parent.numItems; i++) {
            var item = parent.item(i);
            if (item instanceof FolderItem && item.name === name) {
                return item;
            }
        }

        return parent.items.addFolder(name);
    }

    function getFileExtension(filename) {
        if (!filename) return "";
        var lastDot = filename.lastIndexOf(".");
        if (lastDot === -1) return "";
        return filename.substring(lastDot).toLowerCase();
    }

    function isVideoFile(filename) {
        var ext = getFileExtension(filename);
        for (var i = 0; i < config.videoExts.length; i++) {
            if (ext === config.videoExts[i]) return true;
        }
        return false;
    }

    function isImageFile(filename) {
        var ext = getFileExtension(filename);
        for (var i = 0; i < config.imageExts.length; i++) {
            if (ext === config.imageExts[i]) return true;
        }
        return false;
    }

    function isAudioFile(filename) {
        var ext = getFileExtension(filename);
        for (var i = 0; i < config.audioExts.length; i++) {
            if (ext === config.audioExts[i]) return true;
        }
        return false;
    }

    function isCompUsed(comp) {
        for (var i = 1; i <= proj.numItems; i++) {
            var item = proj.item(i);
            if (item instanceof CompItem && item !== comp) {
                for (var j = 1; j <= item.numLayers; j++) {
                    if (item.layer(j).source === comp) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    function getCompUsageInfo(comp) {
        var usedIn = [];
        for (var i = 1; i <= proj.numItems; i++) {
            var item = proj.item(i);
            if (item instanceof CompItem && item !== comp) {
                for (var j = 1; j <= item.numLayers; j++) {
                    if (item.layer(j).source === comp) {
                        usedIn.push(item.name);
                        break;
                    }
                }
            }
        }
        return usedIn;
    }

    function getFootageInComp(comp) {
        var footageList = [];
        for (var i = 1; i <= comp.numLayers; i++) {
            var layer = comp.layer(i);
            if (layer.source && layer.source instanceof FootageItem) {
                var found = false;
                for (var j = 0; j < footageList.length; j++) {
                    if (footageList[j] === layer.source) {
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    footageList.push(layer.source);
                }
            }
        }
        return footageList;
    }



    var mainFolders = {
        videos: getOrCreateFolder(config.folders.videos),
        images: getOrCreateFolder(config.folders.images),
        audio: getOrCreateFolder(config.folders.audio),
        compositions: getOrCreateFolder(config.folders.compositions),
        solids: getOrCreateFolder(config.folders.solids),
        footage: getOrCreateFolder(config.folders.footage),
        other: getOrCreateFolder(config.folders.other)
    };



    var items = {
        videos: [],
        images: [],
        audio: [],
        compositions: [],
        precomps: [],
        unusedComps: [],
        solids: [],
        other: []
    };

    for (var i = 1; i <= proj.numItems; i++) {
        var item = proj.item(i);

        if (item instanceof FolderItem) {
            continue;
        }

        if (item instanceof CompItem) {
            items.compositions.push(item);
        } else if (item instanceof FootageItem) {
            if (item.mainSource instanceof SolidSource) {
                items.solids.push(item);
            } else if (item.file) {
                var filename = item.file.name;
                if (isVideoFile(filename)) {
                    items.videos.push(item);
                } else if (isImageFile(filename)) {
                    items.images.push(item);
                } else if (isAudioFile(filename)) {
                    items.audio.push(item);
                } else {
                    items.other.push(item);
                }
            } else {
                items.other.push(item);
            }
        }
    }


    for (var i = 0; i < items.compositions.length; i++) {
        var comp = items.compositions[i];
        if (isCompUsed(comp)) {
            items.precomps.push(comp);
        } else {
            items.unusedComps.push(comp);
        }
    }



    for (var i = 0; i < items.videos.length; i++) {
        var video = items.videos[i];
        var baseName = video.name.replace(/\.[^\.]+$/, ""); // Retirer l'extension
        var folder = getOrCreateFolder(baseName, mainFolders.videos);
        video.parentFolder = folder;
    }



    for (var i = 0; i < items.images.length; i++) {
        var image = items.images[i];
        var baseName = image.name.replace(/\.[^\.]+$/, "");

        // Grouper les séquences d'images
        if (image.mainSource instanceof FileSource && image.mainSource.isStill === false) {
            var folder = getOrCreateFolder("Séquence_" + baseName, mainFolders.images);
            image.parentFolder = folder;
        } else {
            var folder = getOrCreateFolder(baseName, mainFolders.images);
            image.parentFolder = folder;
        }
    }



    for (var i = 0; i < items.audio.length; i++) {
        var audio = items.audio[i];
        var baseName = audio.name.replace(/\.[^\.]+$/, "");
        var folder = getOrCreateFolder(baseName, mainFolders.audio);
        audio.parentFolder = folder;
    }



    for (var i = 0; i < items.solids.length; i++) {
        items.solids[i].parentFolder = mainFolders.solids;
    }



    var precompsFolder = getOrCreateFolder(config.folders.precomps, mainFolders.compositions);

    for (var i = 0; i < items.precomps.length; i++) {
        var comp = items.precomps[i];
        var usageInfo = getCompUsageInfo(comp);

        if (usageInfo.length > 0) {
            var parentCompName = usageInfo[0].replace(/[\/\\:*?"<>|]/g, "_");
            var folder = getOrCreateFolder(parentCompName, precompsFolder);
            comp.parentFolder = folder;
        } else {
            comp.parentFolder = precompsFolder;
        }
    }

    if (items.unusedComps.length > 0) {
        var unusedFolder = getOrCreateFolder(config.folders.unusedComps, mainFolders.compositions);
        for (var i = 0; i < items.unusedComps.length; i++) {
            items.unusedComps[i].parentFolder = unusedFolder;
        }
    }

    var mainCompsFolder = getOrCreateFolder("Main Compositions", mainFolders.compositions);

    for (var i = 0; i < items.compositions.length; i++) {
        var comp = items.compositions[i];

        var isPrecomp = false;
        for (var j = 0; j < items.precomps.length; j++) {
            if (items.precomps[j] === comp) {
                isPrecomp = true;
                break;
            }
        }

        var isUnused = false;
        for (var j = 0; j < items.unusedComps.length; j++) {
            if (items.unusedComps[j] === comp) {
                isUnused = true;
                break;
            }
        }

        if (!isPrecomp && !isUnused) {
            var footage = getFootageInComp(comp);

            if (footage.length > 0) {
                var footageName = footage[0].name.replace(/\.[^\.]+$/, "").replace(/[\/\\:*?"<>|]/g, "_");
                var folder = getOrCreateFolder(footageName, mainCompsFolder);
                comp.parentFolder = folder;
            } else {
                comp.parentFolder = mainCompsFolder;
            }
        }
    }


    for (var i = 0; i < items.other.length; i++) {
        items.other[i].parentFolder = mainFolders.other;
    }

    var report = "✅ ORGANISATION TERMINÉE !\n\n";
    report += "- Videos : " + items.videos.length + "\n";
    report += "- Images : " + items.images.length + "\n";
    report += "- Audio : " + items.audio.length + "\n";
    report += "- Main compositions : " + (items.compositions.length - items.precomps.length - items.unusedComps.length) + "\n";
    report += "-  : " + items.precomps.length + "\n";
    report += "- Unused compositions : " + items.unusedComps.length + "\n";
    report += "- Solids : " + items.solids.length + "\n";
    report += "- Others : " + items.other.length + "\n";

    alert(report);

    app.endUndoGroup();

})();