$(function () {
	$('#traveldate').datetimepicker({ format: 'DD/MM/YYYY'}).on("dp.change", function (e) {
        $("#midbar").formValidation('revalidateField', 'traveldate');
     });
	$("#midbar").formValidation({
		framework: "bootstrap",
		excluded: ':disabled',
		icon: {
			valid: "glyphicon glyphicon-ok",
			invalid: "glyphicon glyphicon-remove",
			validating: "glyphicon glyphicon-refresh"
		},
		fields: {
			traveldate: {row:'.col-md-6',validators:{notEmpty: {message: "The Travel Date is required"},
			date:{format: 'DD/MM/YYYY',message: 'The Travel Date is not a valid date'}}},
			destination:{row:'.col-md-6',validators:{notEmpty:{message: "The destination is required"}}},
			residence:{row:'.col-md-6',validators:{notEmpty:{message:"The residence is required"},}},
		}
	}).on('success.form.fv', function(e) {
        e.preventDefault();
        var data = [];
        var fileArray = [];
        
        $("#attachFilesContainer input:file").each(function () {
        	if ($(this)[0].files[0]) {                    
        		fileArray.push({ "Attachment": $(this)[0].files[0] });                    
        	}
        });
        var status="Pending";
        if($("#reqstatus").val() =="Draft") status="Draft"
        data.push({"Status":status,"Origin": $("#residence").val(),"Department":awf_user.department,"Office":awf_user.office,"Destination": $("#destination").val(),"DateofTravel": moment($('#trdate').val(), 'DD/MM/YYYY').format('YYYY-MM-DDThh:mm:ss[Z]'), "Files": fileArray});            
        createItemWithAttachments("Request", data).then(
        	function(){
        	   $('#midbar').formValidation('resetForm', true);	
        		swal('success','Visa request submitted succesfully','success');
        		
        	},
        	function(sender, args){
        		swal('Error','Error occured' + args.get_message());
        	}
        	)

    });
	
	$("#btnSubmit").click(function(){ $("#reqstatus").val("Publish"); $("#midbar").data('formValidation').validate()});
	$("#btnDraft").click(function(){ $("#reqstatus").val("Draft"); $("#midbar").data('formValidation').validate()});

	
	var createItemWithAttachments = function(listName, listValues){			
		var fileCountCheck = 0;
		var fileNames;			
		var context = new SP.ClientContext.get_current();
		var dfd = $.Deferred();
		var targetList = context.get_web().get_lists().getByTitle(listName);
		context.load(targetList);
		var itemCreateInfo = new SP.ListItemCreationInformation();
		var listItem = targetList.addItem(itemCreateInfo);
		listItem.set_item("Status", listValues[0].Status);	 
		listItem.set_item("Origin", listValues[0].Origin);
		listItem.set_item("Destination", listValues[0].Destination); 
		listItem.set_item("DateofTravel", listValues[0].DateofTravel); 
		listItem.set_item("Department", listValues[0].Department);
		listItem.set_item("Office", listValues[0].Office); 
		listItem.update();
		context.executeQueryAsync(
			function () {
				var id = listItem.get_id();
				if (listValues[0].Files.length != 0) {
					if (fileCountCheck <= listValues[0].Files.length - 1) {
						loopFileUpload(listName, id, listValues, fileCountCheck).then(
							function () {
							},
							function (sender, args) {
								console.log("Error uploading");
								dfd.reject(sender, args);
							});
					}
				}
				else {
					dfd.resolve(fileCountCheck);
				}
			},   
			function(sender, args){
				swal('Error','Error occured' + args.get_message(),'error');	        	
			}
			);
		return dfd.promise();			
	}
	
});	

	function loopFileUpload(listName, id, listValues, fileCountCheck) {
		var dfd = $.Deferred();
		uploadFile(listName, id, listValues[0].Files[fileCountCheck].Attachment).then(
			function (data) {	                   
				var objcontext = new SP.ClientContext();
				var targetList = objcontext.get_web().get_lists().getByTitle(listName);
				var listItem = targetList.getItemById(id);
				objcontext.load(listItem);
				objcontext.executeQueryAsync(function () {
					console.log("Reload List Item- Success");	                                     
					fileCountCheck++;
					if (fileCountCheck <= listValues[0].Files.length - 1) {
						loopFileUpload(listName, id, listValues, fileCountCheck);
					} else {
						swal(fileCountCheck +' file(s) uploaded successfully');
						$('#midbar').formValidation('resetForm', true);
						$("#StatusModal").modal('hide');
						$("#statusDate").val(""); 
	        			AjaxReload(queryuserRequests).success(function(data){ getRequests(data.value)});
					}
				},
				function (sender, args) {
					swal('Error','Error occured' + args.get_message(),'error');	
				});	                 
				
			},
			function (sender, args) {
				console.log("Not uploaded");
				dfd.reject(sender, args);
			}
			);
		return dfd.promise();
	}
	
	function uploadFile(listName, id, file) {
		var deferred = $.Deferred();
		var fileName = file.name;
		getFileBuffer(file).then(
			function (buffer) {
				var bytes = new Uint8Array(buffer);
				var binary = '';
				for (var b = 0; b < bytes.length; b++) {
					binary += String.fromCharCode(bytes[b]);
				}
				var scriptbase = _spPageContextInfo.webServerRelativeUrl + "/_layouts/15/";
				console.log(' File size:' + bytes.length);
				$.getScript(scriptbase + "SP.RequestExecutor.js", function () {
					var createitem = new SP.RequestExecutor(_spPageContextInfo.webServerRelativeUrl);
					createitem.executeAsync({
						url: _spPageContextInfo.webServerRelativeUrl + "/_api/web/lists/GetByTitle('" + listName + "')/items(" + id + ")/AttachmentFiles/add(FileName='" + file.name + "')",
						method: "POST",
						binaryStringRequestBody: true,
						body: binary,
						success: fsucc,
						error: ferr,
						state: "Update"
					});
					function fsucc(data) {
						console.log(data + ' uploaded successfully');
						deferred.resolve(data);
					}
					function ferr(data) {
						console.log(fileName + "not uploaded error");
						deferred.reject(data);
					}
				});

			},
			function (err) {
				deferred.reject(err);
			}
			);
		return deferred.promise();
	}
	function getFileBuffer(file) {
		var deferred = $.Deferred();
		var reader = new FileReader();
		reader.onload = function (e) {
			deferred.resolve(e.target.result);
		}
		reader.onerror = function (e) {
			deferred.reject(e.target.error);
		}
		reader.readAsArrayBuffer(file);
		return deferred.promise();
	}

