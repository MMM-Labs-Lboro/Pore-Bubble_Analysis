// Macro to process multiple Bread images in a folder
// Sub-folders allowed but will be ignored

#@ File (label = "Input directory", style = "directory") input
#@ String (choices={"Prep Rustic", "Rename Rings"}, style="radioButtonHorizontal", persist=false) operation
output = "C:/";
processFolder(input);

// function to scan folders to find files with correct suffix  //Adds Subfolders
function processFolder(input) {
	list = getFileList(input);
	list = Array.sort(list);
	for (i = 0; i < list.length; i++) {
		if(File.isDirectory(input + list[i])) //File.separator + list[i]))
			processFolder(input + list[i]); //File.separator + list[i]);
		if(endsWith(list[i], ".tif"))
			processFile(input, output, list[i]);
	}
}

function processFile(input, output, file) {
	if (operation == "Rename Rings"){
	path = input+File.separator+"Renamed"+File.separator;
	File.makeDirectory(path);
	open(input+File.separator+file);
	run("Rename Macro");
	}
	if (operation == "Prep Rustic"){
	path = input+File.separator+"Prepped"+File.separator;
	File.makeDirectory(path);
	open(input+File.separator+file);
	run("Prep Image Rustic");
	}
}

print("Done");