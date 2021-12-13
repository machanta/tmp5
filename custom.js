

/**
 * Resize function without multiple trigger
 * 
 * Usage:
 * $(window).smartresize(function(){  
 *     // code here
 * });
 */
(function($,sr){
    // debouncing function from John Hann
    // http://unscriptable.com/index.php/2009/03/20/debouncing-javascript-methods/
    var debounce = function (func, threshold, execAsap) {
      var timeout;

        return function debounced () {
            var obj = this, args = arguments;
            function delayed () {
                if (!execAsap)
                    func.apply(obj, args); 
                timeout = null; 
            }

            if (timeout)
                clearTimeout(timeout);
            else if (execAsap)
                func.apply(obj, args);

            timeout = setTimeout(delayed, threshold || 100); 
        };
    };

    // smartresize 
    jQuery.fn[sr] = function(fn){  return fn ? this.bind('resize', debounce(fn)) : this.trigger(sr); };

})(jQuery,'smartresize');
/**
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

var CURRENT_URL = window.location.href.split('#')[0].split('?')[0],
    $BODY = $('body'),
    $MENU_TOGGLE = $('#menu_toggle'),
    $SIDEBAR_MENU = $('#sidebar-menu'),
    $SIDEBAR_FOOTER = $('.sidebar-footer'),
    $LEFT_COL = $('.left_col'),
    $RIGHT_COL = $('.right_col'),
    $NAV_MENU = $('.nav_menu'),
    $FOOTER = $('footer'),
    $DASHBOARD_MENU = $('#dashboard-menu1');
    $PERSONA_MENU = $('#dashboard-menu2');
    $BPMN_MENU = $('#bpmn-menu');
    $RACI_MENU = $('#raci-menu');

// Sidebar
$(document).ready(function() {
    // TODO: This is some kind of easy fix, maybe we can improve this
    var setContentHeight = function () {
        // reset height
    /*    $RIGHT_COL.css('min-height', $(window).height());

        var bodyHeight = $BODY.outerHeight(),
            footerHeight = $BODY.hasClass('footer_fixed') ? -10 : $FOOTER.height(),
            leftColHeight = $LEFT_COL.eq(1).height() + $SIDEBAR_FOOTER.height(),
            contentHeight = bodyHeight < leftColHeight ? leftColHeight : bodyHeight;

        // normalize content
        contentHeight -= $NAV_MENU.height() + footerHeight;

        $RIGHT_COL.css('min-height', contentHeight);

        */
    };

 var opsNew = false;   
 var issues = null;
 var loadIssues = function(data) {
    //RED.opsProjects = data;
   // console.log(JSON.stringify(data));
   issues = data;
  //  data.forEach(function(o){
            
   // });
}
// Load Ops Projects
var loadOpsProjects = function(data) {
    //RED.opsProjects = data;
    console.log(JSON.stringify(data));
    data.forEach(function(o){
             $('#fmi-ops-project-select').append(new Option(o._id, o._id));
    });
}
var loadOpsModel=function(data) {
  if (data) {
    opsNew = false;
    loadModel(data[0].data, data[0]._id, "User1");
  }
}

var issueCreated=function(data) {
    if (data) {
        getData('/jira/getIssues', function(data) { 
            issues = data;
            $('#fmi-show-issue').hide();
            $('#fmi-start-issue').hide();
        });
    }
};
$('#fmi-jira-start-btn').on('click', function(){
    $('fmi-jira-start-btn').prop("disabled",true);
    var body = {summary:$('#raci-edit-summary').val(), desc:$('#raci-edit-description').val()};
     
    postData('/jira/createIssue', body, issueCreated);
});
$('#fmi-ops-project-select').on('change', function(){
   var val =  $('#fmi-ops-project-select').val() ;
   if (val != 'None') {
    getData('/system/getFullOpsModels?_id='+val, loadOpsModel);
   }
});
    let upload = document.getElementById("import-raci-file");
   /* var options = {
        container: 'luckysheet', //luckysheet is the container id
        showinfobar:false,
    }
    luckysheet.create(options); */
    getData('/jira/getIssues', loadIssues);
    getData('/system/getOpsModels', loadOpsProjects);
    if(upload){
        upload.addEventListener("change", function(evt){
            var files = evt.target.files;
            if(files==null || files.length==0){
                alert("No files wait for import");
                return;
            }
            opsNew = true;
            let name = files[0].name;
            let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length-1];
            if(suffix!="xlsx"){
                alert("Currently only supports the import of xlsx files");
                return;
            }

            LuckyExcel.transformExcelToLucky(files[0], function(exportJson, luckysheetfile){
                
                if(exportJson.sheets==null || exportJson.sheets.length==0){
                    alert("Failed to read the content of the excel file, currently does not support xls files!");
                    return;
                }
                loadModel(exportJson.sheets, exportJson.info.name,exportJson.info.name.creator );
            });
        });
    }


var loadModel = function(sheets, name,user) {

    getData('/jira/getProjects', loadProjects);
    
    
    $("#fmi-jira-project").removeClass('hide');
    $('#fmi-jira-project').show();
    $("#fmi-save-model").removeClass('hide');
    $('#fmi-save-model').show();
    $('#fmi-jira-side').show();
    $('#save-raci-status').text('');
    $('#save-raci-btn').prop("disabled",false);

    $('#fmi-jira-tabs li a:not(:first)').addClass('inactive');
    $('.fmi-jira-container').hide();
    $('.fmi-jira-container:first').show();
    $('#fmi-jira-tabs li a').click(function(){
        var t = $(this).attr('id');
      if($(this).hasClass('inactive')){ //this is the start of our condition 
        $('#fmi-jira-tabs li a').addClass('inactive');           
        $(this).removeClass('inactive');
        
        $('.fmi-jira-container').hide();
        $('#fmi-jira-'+ t + '-cont').fadeIn('slow');
     }
    });
    window.luckysheet.destroy();
                

    window.luckysheet.create({
        hook: {
            
       cellMousedown : function(cell,position,sheet,context) {
            if (position.c == 0 ) {
                $('#fmi-show-issue').hide();
                $('#fmi-start-issue').hide();
                $('#fmi-selected-cell').text(cell.m);
                var matchFound = false;
                var matched = null;
                issues.issues.forEach(function(issue){
                    console.log(issue.summary + ' ## ' + cell.m);
                          if (issue.fields.summary == cell.m) {
                              matchFound = true;
                              matched = issue;
                          }
                });
                $('#raci-edit-description').val('');
                if (matchFound) {
                    console.log(JSON.stringify(matched));
                    $('#fmi-show-issue').show();
                    $('#x-c1').text(matched.key);
                    $('#x-c2').text(matched.fields.status.name);
                    $('#x-c3').text(matched.fields.summary);
                    $('#x-c4').text(matched.fields.description.content[0].content[0].text);
                   
                    $('#x-c5').text(matched.fields.assignee.displayName);
                    $('#x-c6').text(matched.fields.created);
                   
                   

                } else {
                    $('#fmi-start-issue').show();
                    $('#raci-edit-summary').val(cell.m);
                    console.log('No Match found');
                }
                return false;
            }
       }  , 
       cellUpdated: function(r,c,old,newV) {
         if (isInRange(r,c)) {  
            if (newV && newV.m == 'R' ) {
              //  if (isUniqueInRow(r,c,old,newV))
                  window.luckysheet.setCellFormat(r,c,"bg", "#3d85c6");
             //     else {
              //        alert("The Value " + newV.m + " should be unique in row. Please change the other value first");
               //       window.luckysheet.setCellValue(r,c,{v:old.v, m:old.m,bg:'#ff0000'});
              //    }
            } else
            if (newV && newV.m == 'A' ) {
                if (isUniqueInRow(r,c,old,newV))
                   window.luckysheet.setCellFormat(r,c,"bg", "#cc0000");
                 else {
                    alert("The Value " + newV.m + " should be unique in row. Please change the other value first");
                    window.luckysheet.setCellValue(r,c,{v:old.v, m:old.m,bg:'#ff0000'});
                }                               
            } else
            if (newV && newV.m == 'I' ) {
                window.luckysheet.setCellFormat(r,c,"bg", "#93c47d");
            } else 
            if (newV && newV.m == 'C' ) {
                window.luckysheet.setCellFormat(r,c,"bg", "#ffd966");
            } else 
            if (newV && (newV.m == 'A,R'  || newV.m == 'R,A')) {
                window.luckysheet.setCellFormat(r,c,"bg", "#cc0000");
            } else {
                alert('Valid Values are "R", "A", "C", "I" or "A,R" '  );
                window.luckysheet.setCellValue(r,c,{v:old.v, m:old.m,bg:'#ff0000'});
            
            }
        }                         
            function isInRange(r,c) {
             var range =    window.luckysheet.getRangeWithFlatten([{"row":[1,100], "column":[1,25]}]);
               // console.log(JSON.stringify(range))
               var ret = false;
               var exp = {};
               try {
                if (range ) {
                    range.forEach(function (o){
                        if (o.r == r && o.c == c) {ret = true; throw exp;}
                    });
                }
            } catch (e) {if (e===exp )ret = true; else throw e;}  
                return ret;
            }

            function isUniqueInRow(r,c, old,newV) {
                var range =    window.luckysheet.getRangeWithFlatten([{"row":[r,r], "column":[1,25]}]);
                  // console.log(JSON.stringify(range))
                  var ret = true;
                  var exp = {};
                  try {
                   if (range ) {
                       range.forEach(function (o){
                           if (o.c != c && window.luckysheet.getCellValue(o.r, o.c) == newV.m)
                            { throw exp;}
                       });
                   }
               } catch (e) {if (e===exp ){ret = false;   } else throw e;}  
                   return ret;
               }                        
        }},

        container: 'luckysheet', //luckysheet is the container id
        showinfobar:false,
        data:sheets,
        title:name,
        userInfo:user
    });
    $('#save-raci-btn').on('click', function() {
        $('#save-raci-btn').prop("disabled",true);
         var data =  window.luckysheet.getAllSheets();
         var name = null;
         var body = null;
         if (opsNew) {
           name = $('#import-raci-file-nm').val();
           body = {_id:name, data:data};
           postData('/system/saveOpsModel', body, finishSaved);
           console.log('saving Ops Model');
         }
        else {
           name =  $('#fmi-ops-project-select').val();
           body = {_id:name, data:data};
           postData('/system/updateOpsModel', body, finishSaved);
           console.log('updating Ops Model');
        }
         
         
       //  console.log(JSON.stringify(body));
         
         
    });

} 
    //finish Save
 var  finishSaved = function(data) {
    $('#save-raci-status').text("Model sucessfully saved!");
    console.log("Save Ops Model" + JSON.stringify(data));
    $('#save-raci-btn').prop("disabled",false);
 }  
// Load Jira Projects
var loadProjects = function(data) {
    data.forEach(function(o){
             $('#fmi-jira-project-select').append(new Option(o.name, o.id));
    });
}

// Get Data
function getData(url, done) {
    $.ajax({
        headers: {
            "Accept": "application/json",
            "Cache-Control": "no-cache"
        },
        dataType: "json",
        cache: false,
        url: url,
        success: function (data) {
             
            done(data);
        },
        error: function(jqXHR,textStatus,errorThrown) {
            console.log("Unexpected error loading : " + url ,jqXHR.status,textStatus);
        }
    });
}

// Get Data
function postData(url, body, done) {
 

    $.ajax({
        url:url,
        type: "POST",
        data: JSON.stringify(body),
        contentType: "application/json; charset=utf-8"
    }).done(function(data,textStatus,xhr) {
        done(data);
    }).fail(function(xhr,textStatus,err) {
        console.log("Unexpected error loading : " + url ,jqXHR.status,textStatus);
    });
}


    //$SIDEBAR_MENU.find('a').on('click', function(ev) {

    $("#elixr-logout").on('click', function(ev) {
        ev.preventDefault();
        RED.user.logout();
    });



     
     $DASHBOARD_MENU.on('click', function(ev) {

        // Created Problem with Sidebar closed by Dragging..
        showDashboard(false);
     });
   
     $PERSONA_MENU.on('click', function(ev) {

        // Created Problem with Sidebar closed by Dragging..
        showPersona(false);
     });
     $BPMN_MENU.on('click', function(ev) {

        // Created Problem with Sidebar closed by Dragging..
        showBPMN(false);
     });
     $RACI_MENU.on('click', function(ev) {

        // Created Problem with Sidebar closed by Dragging..
        showRACI(false);
     });

     $SIDEBAR_MENU.find('a').on('click', function(ev) {
        var $li = $(this).parent();

        if ($li.is('.active')) {
            $li.removeClass('active active-sm');
            $('ul:first', $li).slideUp(function() {
                setContentHeight();
            });
        } else {
            // prevent closing menu if we are on child menu
            if (!$li.parent().is('.child_menu')) {
                $SIDEBAR_MENU.find('li').removeClass('active active-sm');
                $SIDEBAR_MENU.find('li ul').slideUp();
            }
             
            $li.addClass('active');

            $('ul:first', $li).slideDown(function() {
                setContentHeight();
            });
        }
    });

    // toggle small or large menu
    $MENU_TOGGLE.on('click', function() {
        if ($BODY.hasClass('nav-md')) {
            $SIDEBAR_MENU.find('li.active ul').hide();
            $SIDEBAR_MENU.find('li.active').addClass('active-sm').removeClass('active');
            $("#workspace").css("left",72);
            $("#dashboard").css("left",72);
            $("#persona").css("left",72);
            $("#bpmn").css("left",72);
            $("#raci").css("left",72);
        } else {
            $SIDEBAR_MENU.find('li.active-sm ul').show();
            $SIDEBAR_MENU.find('li.active-sm').addClass('active').removeClass('active-sm');
            $("#workspace").css("left",232);
            $("#dashboard").css("left",232);
            $("#persona").css("left",232);
            $("#bpmn").css("left",232);
            $("#raci").css("left",232);
        }

        $BODY.toggleClass('nav-md nav-sm');

        setContentHeight();

        $('.dataTable').each ( function () { $(this).dataTable().fnDraw(); });
    });

    // check active menu
    $SIDEBAR_MENU.find('a[href="' + CURRENT_URL + '"]').parent('li').addClass('current-page');

    $SIDEBAR_MENU.find('a').filter(function () {
        return this.href == CURRENT_URL;
    }).parent('li').addClass('current-page').parents('ul').slideDown(function() {
        setContentHeight();
    }).parent().addClass('active');

    // recompute content when resizing
    $(window).smartresize(function(){  
        setContentHeight();
    });

    setContentHeight();

    // fixed sidebar
    if ($.fn.mCustomScrollbar) {
        $('.menu_fixed').mCustomScrollbar({
            autoHideScrollbar: true,
            theme: 'minimal',
            mouseWheel:{ preventDefault: true }
        });
    }
});
// /Sidebar

// Panel toolbox
$(document).ready(function() {
    $('.collapse-link').on('click', function() {
        var $BOX_PANEL = $(this).closest('.x_panel'),
            $ICON = $(this).find('i'),
            $BOX_CONTENT = $BOX_PANEL.find('.x_content');
        
        // fix for some div with hardcoded fix class
        if ($BOX_PANEL.attr('style')) {
            $BOX_CONTENT.slideToggle(200, function(){
                $BOX_PANEL.removeAttr('style');
            });
        } else {
            $BOX_CONTENT.slideToggle(200); 
            $BOX_PANEL.css('height', 'auto');  
        }

        $ICON.toggleClass('fa-chevron-up fa-chevron-down');
    });

    $('.close-link').click(function () {
        var $BOX_PANEL = $(this).closest('.x_panel');

        $BOX_PANEL.remove();
    });
});
// /Panel toolbox

// Tooltip
$(document).ready(function() {
    $('[data-toggle="tooltip"]').tooltip({
        container: 'body'
    });
});
// /Tooltip

// Progressbar
$(document).ready(function() {
	if ($(".progress .progress-bar")[0]) {
	    $('.progress .progress-bar').progressbar();
	}
});
// /Progressbar

// Switchery
$(document).ready(function() {
    if ($(".js-switch")[0]) {
        var elems = Array.prototype.slice.call(document.querySelectorAll('.js-switch'));
        elems.forEach(function (html) {
            var switchery = new Switchery(html, {
                color: '#26B99A'
            });
        });
    }
});
// /Switchery

// iCheck
$(document).ready(function() {
    if ($("input.flat")[0]) {
        $(document).ready(function () {
            $('input.flat').iCheck({
                checkboxClass: 'icheckbox_flat-green',
                radioClass: 'iradio_flat-green'
            });
        });
    }
});



// /iCheck



$(document).ready(function() {
     showDashboard(true);
     RED.events.on("loginDone" , function(){
        showDashboard(true);
     });
});



// Table
$('table input').on('ifChecked', function () {
    checkState = '';
    $(this).parent().parent().parent().addClass('selected');
    countChecked();
});
$('table input').on('ifUnchecked', function () {
    checkState = '';
    $(this).parent().parent().parent().removeClass('selected');
    countChecked();
});

var checkState = '';

$('.bulk_action input').on('ifChecked', function () {
    checkState = '';
    $(this).parent().parent().parent().addClass('selected');
    countChecked();
});
$('.bulk_action input').on('ifUnchecked', function () {
    checkState = '';
    $(this).parent().parent().parent().removeClass('selected');
    countChecked();
});
$('.bulk_action input#check-all').on('ifChecked', function () {
    checkState = 'all';
    countChecked();
});
$('.bulk_action input#check-all').on('ifUnchecked', function () {
    checkState = 'none';
    countChecked();
});

function countChecked() {
    if (checkState === 'all') {
        $(".bulk_action input[name='table_records']").iCheck('check');
    }
    if (checkState === 'none') {
        $(".bulk_action input[name='table_records']").iCheck('uncheck');
    }

    var checkCount = $(".bulk_action input[name='table_records']:checked").length;

    if (checkCount) {
        $('.column-title').hide();
        $('.bulk-actions').show();
        $('.action-cnt').html(checkCount + ' Records Selected');
    } else {
        $('.column-title').show();
        $('.bulk-actions').hide();
    }
}

// Accordion
$(document).ready(function() {
    $(".expand").on("click", function () {
        $(this).next().slideToggle(200);
        $expand = $(this).find(">:first-child");

        if ($expand.text() == "+") {
            $expand.text("-");
        } else {
            $expand.text("+");
        }
    });

   // init_sparklines();
});

// NProgress
if (typeof NProgress != 'undefined') {
    $(document).ready(function () {
        NProgress.start();
    });

    $(window).on('load', function() {
        NProgress.done();
    });
}

function showRACI(hideSideBar) {
    hideAll(hideSideBar);
    $("#raci").removeClass('hide');
    $("#raci").show();
    $('#luckysheet').show();

}
function showBPMN(hideSideBar) {
   
    hideAll(hideSideBar);
    $("#bpmn").removeClass('hide');
    $("#bpmn").show();

}
function showPersona(hideSideBar) {
    hideAll(hideSideBar);
    $("#persona").removeClass('hide');
    $("#persona").show();

}

function showDashboard(hideSideBar) {

    hideAll(hideSideBar);
    $("#dashboard").removeClass('hide');
    $("#dashboard").show();
  
    initSpark();
     init_chart_doughnut();   
     init_flot_chart();
     init_gauge();
}
function hideAll(hideSideBar) {
    $("#workspace").hide();
    $("#editor-stack").hide();
    if ( !$("#main-container").hasClass("sidebar-closed") ) {
          $("#sidebar").hide();
          $("#sidebar-separator").hide();
    }

    if (hideSideBar) {
        $("#sidebar").hide();
        $("#sidebar-separator").hide();
    }
    $("#dashboard").hide();
    $("#persona").hide();
    $("#bpmn").hide();
    $("#raci").hide();
    $('#luckysheet').hide();
}
var chartDonut1 = null;
var chartDonut2 = null;
var chartDonut3 = null;
//Extra

function initSpark() {
    $(".sparkline_one").sparkline([2, 4, 3, 4, 5, 4, 5, 4, 3, 4, 5, 6, 4, 5, 6, 3, 5, 4, 5, 4, 5, 4, 3, 4, 5, 6, 7, 5, 4, 3, 5, 6], {
        type: 'bar',
        height: '125',
        barWidth: 13,
        colorMap: {
            '7': '#87a5c3'
        },
        barSpacing: 2,
        barColor: '#416181'
    });

    $(".sparkline22").sparkline([2, 4, 3, 4, 7, 5, 4, 3, 5, 6, 2, 4, 3, 4, 5, 4, 5, 4, 3, 4, 6], {
        type: 'line',
        height: '40',
        width: '200',
        lineColor: '#416181',
        fillColor: '#ffffff',
        lineWidth: 3,
        spotColor: '#34495E',
        minSpotColor: '#34495E'
    });

    $(".sparkline11").sparkline([2, 4, 3, 4, 5, 4, 5, 4, 3, 4, 6, 2, 4, 3, 4, 5, 4, 5, 4, 3], {
        type: 'bar',
        height: '40',
        barWidth: 8,
        colorMap: {
            '7': '#87a5c3'
        },
        barSpacing: 2,
        barColor: '#416181'
    });
}

function init_chart_doughnut(){
				
    if( typeof (Chart) === 'undefined'){ return; }
    
    console.log('init_chart_doughnut');
 
    //if ($('.canvasDoughnut').length){
        
    var chart_doughnut_settings1 = {
            type: 'doughnut',
            tooltipFillColor: "rgba(51, 51, 51, 0.55)",
            data: {
                labels: [
                    "AWS",
                    "Storage",
                    "Function",
                    "Other",
                    "Network"
                ],
                datasets: [{
                    data: [15, 20, 30, 10, 30],
                    backgroundColor: [
                        "#020405",
                        "#192632",
                        "#2a3f54",
                        "#46698d",
                        "#a9bed4"
                    ],
                    hoverBackgroundColor: [
                        "#020405",
                        "#192632",
                        "#2a3f54",
                        "#46698d",
                        "#a9bed4"
                    ]
                }]
            },
            options: { 
                 /*legend: {display: true, position:"left",
                    labels: {
                        fontColor: 'rgb(255, 99, 132)',
                        fontSize : 10,
                        padding:7
                    }},    */
                 legend:false,
                responsive: false 
            }
        }
    

        var chart_doughnut_settings2 = {
            type: 'doughnut',
            tooltipFillColor: "rgba(51, 51, 51, 0.55)",
            data: {
                labels: [
                    "Alexa",
                    "Storage",
                    "Function",
                    "Other" 
                    
                ],
                datasets: [{
                    data: [7, 7, 2, 6],
                    backgroundColor: [
                        "#020405",
                        
                        "#2a3f54",
                        "#46698d",
                        "#a9bed4"
                    ],
                    hoverBackgroundColor: [
                        "#020405",
                        
                        "#2a3f54",
                        "#46698d",
                        "#a9bed4"
                    ]
                }]
            },
            options: { 
                 
                 legend:false,
                responsive: false 
            }
        }


        var chart_doughnut_settings3 = {
            type: 'doughnut',
            tooltipFillColor: "rgba(51, 51, 51, 0.55)",
            data: {
                labels: [
                    "AWS",
                    "Storage",
                    "Alexa" 
                    
                    
                ],
                datasets: [{
                    data: [3, 5, 5],
                    backgroundColor: [
                        "#020405",
                        
                        "#46698d",
                        "#a9bed4"
                    ],
                    hoverBackgroundColor: [
                        "#020405",
                       
                        "#46698d",
                        "#a9bed4"
                    ]
                }]
            },
            options: { 
                 
                 legend:false,
                responsive: false 
            }
        }
         if (chartDonut1) 
            chartDonut1.destroy();
         if (chartDonut2) 
            chartDonut2.destroy();

            if (chartDonut3) 
            chartDonut3.destroy();          
       // $('#canvasDoughnut1').remove();
       // $('#canvasDoughnut2').remove();
  
        var chart_element1 =  $('#canvasDoughnut1');
       // chart_element1.append('<canvas id="canvasDoughnut1" height="110" width="110" style="margin: 5px 10px 10px 0"></canvas>');
        
       chartDonut1 = new Chart( chart_element1, chart_doughnut_settings1);
       //chartDonut1.destroy();
         var chart_element2 =  $('#canvasDoughnut2');
       //  chart_element2.append('<canvas id="canvasDoughnut2" height="110" width="110" style="margin: 5px 10px 10px 0"></canvas>');
          chartDonut2 = new Chart( chart_element2, chart_doughnut_settings2);
          var chart_element3 =  $('#canvasDoughnut3');
          chartDonut3 = new Chart( chart_element3, chart_doughnut_settings3);


         // chartDonut2.destroy();
       /* $('.canvasDoughnut').each(function(){
            
            var chart_element = $(this);

            var chart_doughnut = new Chart( chart_element, chart_doughnut_settings);
            chart_doughnut.destroy();
            var chart_doughnut = new Chart( chart_element, chart_doughnut_settings);
            
        });		 */	
    
   // }  
   
}

var randNum = function() {
    return (Math.floor(Math.random() * (1 + 40 - 20))) + 20;
  };


function init_flot_chart(){
		
    if( typeof ($.plot) === 'undefined'){ return; }
    
    console.log('init_flot_chart');
    
    
    
    

    
    
    
    var chart_plot_02_data = [];
    
    var dt =  
                Date.today().add({ days: -30 }) 
    
    for (var i = 0; i < 30; i++) {

        
       
      chart_plot_02_data.push([new Date(dt.add({days:1})).getTime(), randNum() + i + i + 10]);
    }
    
    
   
    
    var chart_plot_02_settings = {
        grid: {
            show: true,
            aboveData: true,
            color: "#3f3f3f",
            labelMargin: 10,
            axisMargin: 0,
            borderWidth: 0,
            borderColor: null,
            minBorderMargin: 5,
            clickable: true,
            hoverable: true,
            autoHighlight: true,
            mouseActiveRadius: 100
        },
        series: {
            lines: {
                show: true,
                fill: true,
                lineWidth: 2,
                steps: false
            },
            points: {
                show: true,
                radius: 4.5,
                symbol: "circle",
                lineWidth: 3.0
            }
        },
        legend: {
            position: "ne",
            margin: [0, -25],
            noColumns: 0,
            labelBoxBorderColor: null,
            labelFormatter: function(label, series) {
                return label + '&nbsp;&nbsp;';
            },
            width: 40,
            height: 1
        },
        colors: ['#35506b', '#3F97EB', '#72c380', '#6f7a8a', '#f7cb38', '#5a8022', '#2c7282'],
        shadowSize: 0,
        tooltip: true,
        tooltipOpts: {
            content: "%s: %y.0",
            xDateFormat: "%d/%m",
        shifts: {
            x: -30,
            y: -50
        },
        defaultTheme: false
        },
        yaxis: {
            min: 0
        },
        xaxis: {
            mode: "time",
            minTickSize: [1, "day"],
            timeformat: "%d/%m/%y",
            min: chart_plot_02_data[0][0],
            max: chart_plot_02_data[20][0]
        }
    };	

    
    
    
    
    
    if ($("#chart_plot_02").length){
        console.log('Plot2');
        
        $.plot( $("#chart_plot_02"), 
        [{ 
            label: "Flows Dispached", 
            data: chart_plot_02_data, 
            lines: { 
                //fillColor: "rgba(150, 202, 89, 0.12)" 
                //fillColor: "rgba(53, 80, 107, 0.1)"
                fillColor: "rgba(113, 147, 183, 0.3)"
                
            }, 
            points: { 
                fillColor: "#fff" } 
        }], chart_plot_02_settings);
        
    }
    
   
  
} 



function init_gauge() {
			
    if( typeof (Gauge) === 'undefined'){ return; }
    
    console.log('init_gauge [' + $('.gauge-chart').length + ']');
    
    console.log('init_gauge');
    

      var chart_gauge_settings = {
      lines: 12,
      angle: 0,
      lineWidth: 0.4,
      pointer: {
          length: 0.75,
          strokeWidth: 0.042,
          color: '#1D212A'
      },
      limitMax: 'false',
      colorStart: '#2a3f54',
      colorStop: '#2a3f54',
      strokeColor: '#b4c6d9',
      generateGradient: true
  };
    
    
    if ($('#chart_gauge_01').length){ 
    
        var chart_gauge_01_elem = document.getElementById('chart_gauge_01');
        var chart_gauge_01 = new Gauge(chart_gauge_01_elem).setOptions(chart_gauge_settings);
        
    }	
    if ($('#chart_gauge_02').length){ 
    
        var chart_gauge_02_elem = document.getElementById('chart_gauge_02');
        var chart_gauge_02 = new Gauge(chart_gauge_02_elem).setOptions(chart_gauge_settings);
        
    }	
    
    if ($('#gauge-text1').length){ 
    
        chart_gauge_01.maxValue = 100;
        chart_gauge_01.animationSpeed = 32;
        chart_gauge_01.set(40);
        chart_gauge_01.setTextField(document.getElementById("gauge-text1"));
    
    }
    
    if ($('#gauge-text2').length){ 
    
        chart_gauge_02.maxValue = 100;
        chart_gauge_02.animationSpeed = 32;
        chart_gauge_02.set(97);
        chart_gauge_02.setTextField(document.getElementById("gauge-text2"));
    
    }
    

}   