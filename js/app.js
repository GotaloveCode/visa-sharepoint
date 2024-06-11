var siteUrl = _spPageContextInfo.webAbsoluteUrl + '/';
var awf_visalst = [],awf_attachments=[],awf_request=[],tbrequest=null,tbreview=null,tbreport=null,awf_request_array=[],awf_origins=[],awf_destinations=[],
awf_user={admin:false,department:null};
var lstUrl = siteUrl  + '_api/web/lists/getbytitle'; 
var queryallRequests="('Request')/items?$select=Id,Department,Office,Comments,Destination, Origin,Status,DateofTravel,Created,Author/Id,Author/Title,AttachmentFiles,AttachmentFiles/ServerRelativeUrl,AttachmentFiles/FileName&$expand=Author,AttachmentFiles&$filter=Status ne 'Draft'";
var queryuserRequests="('Request')/items?$select=Id,Department,Office,Comments,Destination, Origin,Status,DateofTravel,Created,Author/Id,Author/Title,AttachmentFiles,AttachmentFiles/ServerRelativeUrl,AttachmentFiles/FileName&$expand=Author,AttachmentFiles&$filter=Author/Id eq " + _spPageContextInfo.userId;
var queryUser="('User')/items?$select=User/Title,User/Id,Department,Country,Office&$expand=User&$filter=User/Id eq " + _spPageContextInfo.userId;
var queryvisaInfo="('VISA Info')/items?$select=Destination,Origin,Info";

function loadBatch(){
    var commands = [];
    var batchExecutor = new RestBatchExecutor(_spPageContextInfo.webAbsoluteUrl, {'X-RequestDigest': $('#__REQUESTDIGEST').val()});
    batchRequest = new BatchRequest();
    batchRequest.endpoint = lstUrl + queryvisaInfo;
    batchRequest.headers = {'accept': 'application/json;odata=nometadata'}
    commands.push({id: batchExecutor.loadRequest(batchRequest),title: "getVISAInfo"});

    batchRequest = new BatchRequest();
    batchRequest.endpoint = lstUrl + queryuserRequests;
    batchRequest.headers = {'accept': 'application/json;odata=nometadata'}
    commands.push({id: batchExecutor.loadRequest(batchRequest),title: "getRequests"});
    
    batchRequest = new BatchRequest();
    batchRequest.endpoint = lstUrl + queryallRequests;
    batchRequest.headers = {'accept': 'application/json;odata=nometadata'}
    if(awf_user.admin)
        commands.push({id: batchExecutor.loadRequest(batchRequest),title: "getAllrequests"});
    
    batchRequest = new BatchRequest();
    batchRequest.endpoint = lstUrl + queryUser;
    batchRequest.headers = {'accept': 'application/json;odata=nometadata' }
    commands.push({ id: batchExecutor.loadRequest(batchRequest),title: "getUser"});

    batchExecutor.executeAsync().done(function(result) {
        $.each(result, function(k, v) {
            var command = $.grep(commands, function(command) {
                return v.id === command.id;
            });
            if (command[0].title == "getVISAInfo") {
                getVISAInfo(v.result.result.value);
            } else if (command[0].title == "getRequests") {
                getRequests(v.result.result.value);
            } else if (command[0].title == "getUser") {
                getUser(v.result.result.value);
            } else if (command[0].title == "getAllrequests") {
                getAllrequests(v.result.result.value);
            }
        });
    }).fail(function(err) {
        onError(err);
    });
}

function onError(e) {
    swal("Error", e.responseText, "error");
}

// check user has permission to access
function IsCurrentUserHasAdminPerms(){ 
    IsCurrentUserMember("Team Site Owners",function(isCurrentUserInGroup){
        if(isCurrentUserInGroup){
            $('[href="#Reviewrequest"],[href="#Reports"],[href="#admin"]').closest('li').show();
            awf_user.admin= true;                
        }
        loadBatch();
    }
    );
}       
//check user membership that works with nested groups
function IsCurrentUserMember(groupname,OnComplete){
    $.ajax({
     url: _spPageContextInfo.webAbsoluteUrl+"/_api/web/sitegroups/getbyname('"+groupname+"')/CanCurrentUserViewMembership",
     method: "GET",
     headers: {"Accept": "application/json; odata=verbose"},
     success: function(data) {OnComplete(data.d.CanCurrentUserViewMembership)},
     error: function(data) {OnComplete(false)}
 }); 
}

function AjaxReload(query) {
  return $.ajax({
    url: lstUrl+query,
    type: "GET",  
    headers: {'accept': 'application/json;odata=nometadata'},
});
}

/*****update List Item via SharePoint REST interface ******/
function updateJson(endpointUri, payload, success, error) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    $.ajax({
        url: endpointUri,
        type: "POST",
        data: JSON.stringify(payload),
        contentType: "application/json;odata=verbose",
        headers: {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "X-HTTP-Method": "MERGE",
            "If-Match": "*"
        },
        success: success,
        error: onError
    });
}

//post to list function
function postJson(endpointUri, payload, success, error) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    $.ajax({
        url: endpointUri,
        type: "POST",
        data: JSON.stringify(payload),
        contentType: "application/json;odata=verbose",
        headers: {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
        },
        success: success,
        error: onError
    });
}




function getVISAInfo(d){
    var citizenrow = "",destrow = "",deptrow = "";
    $.each(d, function(key, value) {
        awf_visalst.push({
            origin: value.Origin,
            destination: value.Destination,
            info: value.Info
        })
    });

    awf_destinations = alasql('SELECT DISTINCT destination FROM ?', [awf_visalst]);
    if (awf_destinations[0] == "" && awf_destinations.size() > 1)
        awf_destinations = awf_destinations.shift();
    $.each(awf_destinations, function(key, value) {
        destrow += "<option>" + value.destination + "</option>";
    });

    awf_origins = alasql('SELECT DISTINCT origin FROM ?', [awf_visalst]);
    if (awf_origins[0] == "" && awf_origins.size() > 1)
        awf_origins = awf_origins.shift();
    $.each(awf_origins, function(key, value) {
        citizenrow += "<option>" + value.origin + "</option>";
    });

    $("#residence,#selorigin").html(citizenrow).chosen({width:"100%"});
    $("#destination,#seldestination,#currentdest").html(destrow).chosen({width:"100%"});

}

function filterInfo() {
    if ($("#destination").val().trim() != "" && $("#residence").val().trim() != "") {
        var res = alasql('SELECT info FROM ? WHERE destination="' + $("#destination").val() + '" AND origin="' + $("#residence").val() + '"', [awf_visalst]);
        if (res[0]["info"])
            $("#info").html(res[0]["info"]);
    }
}

function getAllrequests(data){
    var row2 ="",row="",opt="",o="",op="",authors=[],origins=[],destinations=[];
    $.each(data, function (key, value) { 
      var created = value.Author.Title;
        row2+='<tr><td>'+moment(value.Created).format("DD-MM-YYYY")+'</td><td>'+created +'</td><td>'+value.Office+'</td><td>'+value.Origin+'</td><td>'+value.Destination+'</td><td>'+moment(value.DateofTravel).format("DD-MM-YYYY")+'</td><td>'+ getAttachmentLinks(value.AttachmentFiles)+'</td><td>'+value.Status+'</td><td>'+value.Comments+'</td><td class="links" data-attachment="'+ getAttachment(value.AttachmentFiles)+'"><a href="#" data-toggle="modal" data-target="#RequestReviewModal" id ="editlink" class ="editlink" data-requestreviewmodaldata=\'{"officeParam":"'+value.Office+'", "traveldateparam":"'+moment(value.DateofTravel).format("DD-MM-YYYY")+'","citizenParam":"'+value.Origin+'","datecreatedParam":"'+moment(value.Created).format("DD-MM-YYYY")+'","createdbyParam":"'+created+'","id":"'+value.Id+'"}\'>Review</a></td></tr>';       
      row+='<tr><td>'+moment(value.Created).format("DD-MM-YYYY")+'</td><td>'+created +'</td><td>'+value.Origin+'</td><td>'+value.Destination+'</td><td>'+moment(value.DateofTravel).format("DD-MM-YYYY")+'</td><td>'+value.Status+'</td></tr>';       
       authors.push(created);
    });
    $('#TableReviewRequest>tbody').html(row2);
    $('#tbreport>tbody').html(row);
    authors =$.unique(authors); 
    $.each(authors,function(i,j){opt+='<option>'+j+'</option>';});
    $("#selcreatedby").html(opt).chosen({width:"100%"});    
    tbreview = $('#TableReviewRequest').dataTable({responsive:true,order:[[ 5, "desc" ]]});
    tbreport = $('#tbreport').dataTable({responsive:true,order:[[ 3, "desc" ]]});

    $(document).on("click", ".editlink", function (){
        $(".RequestReviewModal-body #createdBy").val($(this).data('requestreviewmodaldata').createdbyParam);
        $(".RequestReviewModal-body #destination1").val($(this).data('requestreviewmodaldata').officeParam);
        $(".RequestReviewModal-body #travelDate").val($(this).data('requestreviewmodaldata').traveldateparam);
        $(".RequestReviewModal-body #citizenship").val($(this).data('requestreviewmodaldata').citizenParam);
        $(".RequestReviewModal-body #dateCreated").val($(this).data('requestreviewmodaldata').datecreatedParam);
        $(".RequestReviewModal-body #requestId").val($(this).data('requestreviewmodaldata').id);
    });
    $(document).on("click", "#btnReview", function(){
       reviewRequest($(".RequestReviewModal-body #requestId").val());
   });
}

function getRequests(data){
    var row="";
    $.each(data, function (key, value){ 
        var btn="";
        if(value.Status!="Approved") btn='<a href="#" data-toggle="modal" data-target="#StatusModal" data-statusmodaldata=\'{"traveldate":"'+moment(value.DateofTravel).format("DD-MM-YYYY")+'","id":"'+value.Id+'","status":"'+value.Status+'","destination":"'+value.Destination+'"}\' id ="updatelink">Update</a>';
        row+='<tr><td>'+moment(value.Created).format("DD-MM-YYYY")+'</td><td>'+value.Destination+'</td><td>'+moment(value.DateofTravel).format("DD-MM-YYYY")+'</td><td>'+value.Status+'</td><td>'+value.Comments+'</td><td>'+ getAttachmentLinks(value.AttachmentFiles)+'</td><td class="links" data-attachment="'+ getAttachment(value.AttachmentFiles)+'">'+ btn+'</td></tr>'; 
    });
    $('#TableRequest>tbody').html(row);
    tbrequest =$('#TableRequest').dataTable({responsive:true});
    $(document).on("click", "#updatelink", function(){
      $("#att").html("");
        $("#currenttravelDate").val($(this).data('statusmodaldata').traveldate);
        var id= $(this).data('statusmodaldata').id;
        $("#statusrequestId").val(id);
        $("#currentdest").val($(this).data('statusmodaldata').destination).trigger("chosen:updated");        
        if($(this).data('statusmodaldata').status == "Pending")
          $("#btndelete").show();
        else
          $("#btndelete").hide();
        var p = $(this).parent().data('attachment'),links =[];        
        if(p.length>0){
            p= p.replace(/'/g, '\"');
          p = JSON.parse(p);
          $.each(p, function (key, value){
              links.push('<span><a href="' + value.url+ '" target="_blank">'+ value.filename+ '</a> &nbsp;&nbsp;<span class="fa fa-trash delatt text-danger" style="cursor:pointer" data-id="'+id+'" data-file="'+value.filename+'" ></span></span>');
          });
        }
        $("#att").html(links.join(', '));
    });
    $(document).on("click", "#btnUpdate", function(){
        updateRequest($("#statusrequestId").val());
    });
    
     $(document).on("click", "#btndelete", function(){
        deleteRequest($("#statusrequestId").val());
    });
    
    $(document).on("click",".delatt",function(){
        deleteAttachment($(this).data('id'),$(this).data('file'),$(this));
    });
}

function deleteAttachment(id,filename,elem){
    swal({
        title: "Delete Attachment",
        text: "Are you sure you want to delete this attachment?",
        type:"warning",
        showCancelButton: true,},
        function(isConfirm){ 
       if (isConfirm){
          elem.parent().remove();
          var Url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('Request')/GetItemById(" + id + ")/AttachmentFiles/getByFileName('" + filename+ "')"; 
          $.ajax({ 
            url: Url, 
            type: 'DELETE', 
            contentType: 'application/json;odata=verbose', 
            headers: { 
                'X-RequestDigest': $('#__REQUESTDIGEST').val(), 
                'X-HTTP-Method': 'DELETE', 
                'Accept': 'application/json;odata=verbose' 
            }, 
            success: function (data) { 
                swal("success","Attachment removed successfully"); 
                AjaxReload(queryuserRequests).success(function(data){
                 $("#TableRequest>tbody").empty();
                 getRequests(data.value);
             });
            }, 
            error: onError
        }); 
          
      }
  });
}

function updateRequest(id){
    if($(".statusmodal-body #statusDate").val()!=""){
        var data = [],fileArray = [];
    $("#attFilesContainer input:file").each(function () {
      if ($(this)[0].files[0]) 
              fileArray.push({ "Attachment": $(this)[0].files[0] });
        });
        data.push({"Files": fileArray});
        var item = {"__metadata": { "type": "SP.Data.REQUESTListItem"},"DateofTravel":moment($(".statusmodal-body #statusDate").val(),"DD/MM/YYYY").format("MM/DD/YYYY"),"Destination":$("#currentdest").val(),"Status":"Pending"};
        updateJson(lstUrl+"('Request')/items(" + id + ")",item,requestApproved,onError);
        function requestApproved(){
          if(fileArray.length>0)
            loopFileUpload("Request", id, data, 0);       
          swal('success','Request updated successfully');
        $("#StatusModal").modal('hide');
      $("#statusDate").val(""); 
        AjaxReload(queryuserRequests).success(function(data){ getRequests(data.value)});
     }
   }else swal("Error!","Select new travel date","error");  
}


function deleteRequest(id){
  swal({
        title: "Delete Request",
        text: "Are you sure you want to delete this request?",
        type:"warning",
        showCancelButton: true,},
        function(isConfirm){ 
       if (isConfirm){
          $.ajax({ 
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('Request')/items(" + id + ")", 
            type: 'DELETE', 
            headers: { 
                'X-RequestDigest': $('#__REQUESTDIGEST').val(), 
                'X-HTTP-Method': 'DELETE', 
                'IF-MATCH': '*' 
            }, 
            success: function (data) { 
                swal("success","Request deleted successfully"); 
                AjaxReload(queryuserRequests).success(function(data){
                 $("#TableRequest>tbody").empty();
                 getRequests(data.value);
                 $("#StatusModal").modal('hide');
             });
             
            }, 
            error: onError
        });           
      }     
  });
}


function reviewRequest(id){
    var statusVal =$(".RequestReviewModal-body #status").val();
    var commentsVal = $(".RequestReviewModal-body #comments").val();
    if(statusVal ||commentsVal){ 
     $('#btnReview').attr('disabled','disabled');  
     var item = {"__metadata": { "type": "SP.Data.REQUESTListItem"},"Status":statusVal,"Comments":commentsVal};
     updateJson(lstUrl+"('Request')/items(" + id + ")",item,requestApproved,onError);
     function requestApproved(){
        swal("Success","Request reviewed successfully","success");  
            AjaxReload(queryallRequests).success(function(data){                
                getAllrequests(data.value);
            });
            $("#comments").val(""); 
            $('#btnReview').removeAttr('disabled');  
        }
    }else swal("Error!","Nothing to update","error");  
}


function getUser(u) {
    $.each(u, function(i, j) {
        awf_user.department = j.Department;
        awf_user.office = j.Office;
        awf_user.country = j.Country;
    });
}

$(document).ready(function() {
    IsCurrentUserHasAdminPerms();
    eventsListen();
    $.fn.dataTable.moment( 'DD-MM-YYYY' );
    $(".date").datetimepicker({ format: 'DD-MM-YYYY' });
    $('#dateto,#datefrom').datetimepicker({ format: 'DD-MM-YYYY' }).on('dp.change', function (e) { tbreport._fnReDraw() });
    $(".chosen").chosen({ width: '100%' });
    $('#selcreatedby,#selstatus,#selorigin,#seldestination').change(function(e){ tbreport._fnReDraw()});  
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    
});

function eventsListen() {
  $("#destination,#residence").on("change", function() {
      filterInfo();
  });
}

function getAttachmentLinks(Attachments){  
    var links =[];
    $.each(Attachments, function(index,value){ 
     links.push('<a href="' + value.ServerRelativeUrl+ '" target="_blank">'+ value.FileName+ '</a>');
 });
    return links.join(', ');
}

function getAttachment(Attachments){  
  awf_attachments=[];
  $.each(Attachments, function(index,value){ 
     awf_attachments.push({url:value.ServerRelativeUrl,filename:value.FileName});
 });
  return JSON.stringify(awf_attachments).replace(/"/g, '\'');
}

$.fn.dataTable.ext.search.push(
    function( settings, data, dataIndex ) {
    if (settings.nTable.id != "tbreport") return true;
        date_from = moment('01-01-1000','DD-MM-YYYY');
        the_date = moment().format('DD-MM-YYYY');
        date_to = moment().endOf('year'); 
        if($('#date-from').val() != "") date_from = moment($('#date-from').val(),'DD-MM-YYYY');          
        if($('#date-to').val() != "") date_to = moment($('#date-to').val(),'DD-MM-YYYY');                          
        if(data[4] != "") the_date = data[4];                
        var loc = moment(the_date,'DD-MM-YYYY');        
        if (loc.isSameOrAfter(date_from) && loc.isSameOrBefore(date_to)) return true; 
        return false;
    }
);

$.fn.dataTable.ext.search.push(
    function( settings, data, dataIndex ) {
    if (settings.nTable.id != "tbreport") return true;
        var value = $('#selcreatedby').val();
        var d = data[1];
        if (value == null)
        {
            return true;
        }
        else if(value.indexOf(d) != -1){
            return true;
        }
});


$.fn.dataTable.ext.search.push(
    function( settings, data, dataIndex ) {
    if (settings.nTable.id != "tbreport") return true;
         value = $('#selorigin').val();
        d = data[2];
        if (value == null)
        {
            return true;
        }
        else if(value.indexOf(d) != -1){
            return true;
        }
 });


$.fn.dataTable.ext.search.push(
    function( settings, data, dataIndex ) {
    if (settings.nTable.id != "tbreport") return true;
        value = $('#seldestination').val();
        d = data[3];
        if (value == null)
        {
            return true;
        }
        else if(value.indexOf(d) != -1){
            return true;
        }
});


$.fn.dataTable.ext.search.push(
    function( settings, data, dataIndex ) {
    if (settings.nTable.id != "tbreport") return true;
        value = $('#selstatus').val();
        d = data[5];
        if (value == null)
        {
            return true;
        }
        else if(value.indexOf(d) != -1){
            return true;
        }
});


function loadDocFrames(url) {
  $('.main-loader').show();
  iframe(_spPageContextInfo.webAbsoluteUrl + url, '#documents-iframe', '560px');
} 
function iframe(url, selector, height) {
  $(selector).empty();
  $('<iframe>', {
    src: url,
    id: 'MainIframe',
    'class': 'MainIframe',
    frameborder: 0,
    height: height,
    scrolling: "no",
    width: '100%'
  }).appendTo(selector);
  $('.MainIframe').load(function () {
    $('.main-loader').hide();
    $('.MainIframe').contents().find('body').addClass('ms-fullscreenmode');
    setTimeout(hideFrame, 3000);       
  });
  function hideFrame(){
    $('.MainIframe').contents().find('.od-SuiteNav,.Files-leftNav,.od-TopBar-header.od-Files-header,.footer').hide().css('display', 'none');
    $('.MainIframe').contents().find('.Files-mainColumn').css('left', '0');
    $('.MainIframe').contents().find('.Files-belowSuiteNav').css('top', '0px');
  }
}

