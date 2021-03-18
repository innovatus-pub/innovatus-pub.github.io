/* Toggle between adding and removing the "responsive" class to topnav when the user clicks on the icon */
function foldmenu() {
    var x = document.getElementById("innovnav");
    if (x.className === "sidebar") {
      x.className += " responsive";
    } else {
      x.className = "sidebar";
    }
  }