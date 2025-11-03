document.addEventListener('DOMContentLoaded', RWDlayoutIni);
window.addEventListener('resize', RWDlayoutIni);

// RWD layout function
function RWDlayoutIni() {
    // Set the height of the sidebar menu to create a scrollbar when expanded
    var docHeight = window.innerHeight;
    var mainheaderwrapperHeight = document.getElementById("mainheaderwrapper").offsetHeight;
    var mainfooterHeight = document.getElementById("mainfooter").offsetHeight;
    var sidebarwrapperH = docHeight - mainheaderwrapperHeight - mainfooterHeight + "px";
    document.querySelector(".metismenubox").style.height = sidebarwrapperH;

    // Determine whether to show the sidebar based on screen width
    var docwidth = window.innerWidth;
    var sidebarwrapper = document.getElementById("sidebar-wrapper");
    if (docwidth < 1100) {
        sidebarwrapper.style.display = 'none';
    } else {
        sidebarwrapper.style.display = 'block';
    }

    // Adjust content width
    var containerochiW = docwidth - 40 + "px";
    var containerochiWs = docwidth - 290 + "px";
    var tboverWElements = document.querySelectorAll(".tboverW");

    if (docwidth < 1650 && docwidth > 1101) {
        tboverWElements.forEach(function(element) {
            element.style.width = containerochiWs;
        });
    } else if (docwidth < 1100) {
        tboverWElements.forEach(function(element) {
            element.style.width = containerochiW;
        });
    } else {
        tboverWElements.forEach(function(element) {
            element.style.width = "inherit";
        });
    }
}

// Toggle sidebar menu
document.getElementById("sidebarToggle").addEventListener("click", function() {
    var sidebarwrapper = document.getElementById("sidebar-wrapper");
    if (sidebarwrapper.style.display === 'block' || sidebarwrapper.style.display === '') {
        sidebarwrapper.style.display = 'none';
    } else {
        sidebarwrapper.style.display = 'block';
    }
});