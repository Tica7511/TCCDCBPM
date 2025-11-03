// JavaScript Document
$("#backToTop").hide();
$("body").on('scroll', function () {
    if ($(this).scrollTop() > 100) {
        $("#backToTop").fadeIn();
    } else {
        $("#backToTop").fadeOut();
    }
});
$("#backToTop").on("click",function(){
    $('html, body').animate({ scrollTop: 0 }, 'slow');
})

