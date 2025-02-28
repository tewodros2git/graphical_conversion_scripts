Process flow
	A.run parse_class_names
	-summary-loops through all the scids and create availble port lookup file
	-output- flat file of availble port names
	
	B.run python 0_run_conversion_main <fileName>
	-summary- runs all the process listed below accordingly and converts the file to svg
	- output- converted svg file
	
			Steps:
		1. flat to json
			-summary - parses the flat file and organizes shapes w/ attributes into json
			-inputs - flat file name
			-outputs - json file with flat file objects parsed out
			
		2. get unique names from node script
			-summary - gets the labelId from svg creator to use to map the flat file objects to the svg elements
			-inputs - flat file name
			-outputs - flat file of distinct obj names for mapping

		3. run connection scripts
			-summary - updates json file with unique name, calculates x,y of each port and finds connecting object
			-inputs - json with objects, flat file with distinct names
			-outputs - json file with connection info

		4. run node script
			-summary - parses flat file out into an svg to be used in visio. includes connection info in custom props
			-inputs - flat file, json file with connection info
			-outputs - svg file

		5. Replace "-" with "_" for shapes within svg

	C.Open svg in visio, move object to fit display and flip objects that are in the wrong direction. import macro     	LineConnector and run macro "RunLines" 
		-inputs - svg file
		-outputs - visio file
		
	D.Open svg in visio, import macro LineTracingInfo and run macro  
		-inputs - svg file
		-outputs - json file
		
	E.Run macro "ReplaceShapesWithStencilObjects" 
		-summary-finds switches and replace them with objects from stencil and copy all the properties to the new object and remove the old. Nb: the line connections with switches will broken so need to manually reconnect them. verify switch rotation for each switch.
		-outputs-visio file
		
	F.export svg from visio
	
	G.run updateHotArrowProps to update the hot arrow property hirarchy 
