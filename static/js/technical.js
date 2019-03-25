    /*Menu*/
    $('.menu .item').click(function () {
        $('.item').removeClass('active');
        $(this).addClass('active');
    });

    /*Mobile*/
    $('.ui.sidebar').sidebar({
        context: $('.ui.pushable.segment'),
        transition: 'overlay'
    }).sidebar('attach events', '#mobile_item');