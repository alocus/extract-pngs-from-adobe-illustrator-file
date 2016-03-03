function exportFileToPNG(dest, artBoardIndex)
{
    var exportOptions = new ExportOptionsPNG24(); // or ExportOptionsPNG24
    var type = ExportType.PNG24; // or ExportType.PNG24
    var file = new File(dest + ".png");

    exportOptions.artBoardClipping = true;
    exportOptions.antiAliasing = true;
    exportOptions.transparency = true;
    exportOptions.qualitySetting = 72;
    exportOptions.saveMultipleArtboards = false;
    exportOptions.verticalScale = 500;
    exportOptions.horizontalScale = 500;
    exportOptions.artboardRange = "" + artBoardIndex;
    app.activeDocument.exportFile( file, type, exportOptions );
}

function execute()
{
    if (app.documents.length == 0)
    {
        alert('No document open', 'Error');
        return;
    }

    if (app.activeDocument.selection.length == 0)
    {
        alert('Nothing selected', 'Error');
        return;
    }

    var selectedStuff = app.activeDocument.selection[0];

    // snap position to pixels
    selectedStuff.position = [ Math.round(selectedStuff.position[0]), Math.round(selectedStuff.position[1]) ];

    // create temporary artboad for exporting
    var docRef = app.activeDocument;
    var rect = selectedStuff.visibleBounds;

    try
    {
        docRef.artboards.add(rect);
    }
    catch(e)
    {
        alert('Could not create Artboard as step of export.', 'Failure');
        return;
    }

    // determine destination
    var destFolder = docRef.path;
    if(destFolder == "")
        destFolder = Folder.selectDialog('Select the folder to export to:');

    if(destFolder)
    {
        try
        {
            exportFileToPNG(destFolder + "/" + docRef.name, docRef.artboards.length);
        }
        catch(e) {}
    }

    // delete temp-artboard
    docRef.artboards.remove(docRef.artboards.length - 1);
}

execute();
