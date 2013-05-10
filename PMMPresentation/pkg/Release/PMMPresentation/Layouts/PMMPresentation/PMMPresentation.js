jQuery(document).ready(function () {
    FixSharePointFunction();

    //To change the content type icon
//    jQuery(jQuery("body")[0]).bind("DOMSubtreeModified",
//	function () {
//	    console.log('change');
//	    var img = jQuery("img[alt='PMM Presentation']");
//	    if (img)
//	        img.attr("src", "/_layouts/images/lg_icxlsx.png");
//	});
    //end

});

function FixSharePointFunction() {
    window.oldFunction = window.STSNavigate2;
    window.STSNavigate2 = function STSNavigate2(event, url) {
        if (url.indexOf("/_layouts/PMMPresentation/NewPMMPresentation.aspx") == -1)
            window.oldFunction(event, url);
        else {
            if (url.indexOf("IsDlg") == -1)
                url += "&IsDlg=1";
            NewItem2(event, url);
        }
    };
}