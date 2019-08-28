'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            $('#webBtn').click(function () { insertHTML(0); });
            $('#fileBtn').click(function () { insertHTML(1); });

        });

    });


    function insertHTML(type) {

        var text;
        // 0: web, other: file
        if (type == 0) {
            text =
                '       <input id="web">' +
                '<button class = "btn" id="insertURL" onclick="insURL();" >Insert URL</button>' +
                '      </br> </br>  '
        } else {

            text = '       <input type="file" id="files" onchange = "handleFileSelect(this.files);" />  '

        };

        $("body").html(text);

    }

})();

function handleFileSelect(f) {

    var reader = new FileReader();
    reader.onload = function () {
        var text = reader.result;
        $('body').html(text);

    };
    reader.readAsText(f[0]);

}


function insURL() {

    var url = document.getElementById("web").value;

    $("body").html("Loading...");

    $.ajax({
        url: url,
        type: "GET",
        dataType: "html"
    }).done(function (data) {
        $("body").html(data);
        $("body link").each(function () {
            var cssLink = $(this).attr('href');
            $("body head").append('  <link href="' + cssLink + '" rel="stylesheet" />');
        });
        $("body script").each(function () {
            var jsLink = $(this).attr('src');
            $("body head").append('<script type="text/javascript" src="' + jsLink + '"></script>')
        });

    }).fail(function (jqXHR, textStatus, errorThrown) {
        $("body").html("Error. File is not loaded");
    });



    


}