'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            try {
                var t1 = Office.context.document.settings.get('elementHTML');
                var t2 = '       <div id="start-menu">  ' +
                    '           <button class="btn" id="webBtn">Website</button>  ' +
                    '           <button class="btn" id="fileBtn">Local file</button>  ' +
                    '   		  ' +
                    '           <div id="start-info">  ' +
                    '               <p> This is a free program which main purpose is to load webpage content into a web-api container. </p>  ' +
                    '               <p> If you want to load a website content, press a <strong> Website </strong> button. </p>  ' +
                    '               <p> If you want to load a local html, press a <strong> Local URL </strong> button. </p>  ' +
                    '               <br>  ' +
                    '               <p> After loading a web-element into this container, it will automatically open anywhere without necessity of loading the web-element again.</p>  ' +
                    '               <br>  ' +
                    '               <p> Created by <strong> sanederchik </strong> </p>  ' +
                    '               <p> More on my <a href="https://github.com/sanederchik"> github page </a> </p>  '  + 
                        '           </div>  '  +
                            '   		  ' +
                            '      </div>  '; 


                if (t1 == null) {

                    $('body').html(t2);

                    $('#webBtn').click(function () { insertHTML(0); });
                    $('#fileBtn').click(function () { insertHTML(1); });

                } else {

                    $('body').html(t1);

                };


            } catch (err) {

                $('body').html(t2);
                $('#webBtn').click(function () { insertHTML(0); });
                $('#fileBtn').click(function () { insertHTML(1); });

            }

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

        $('body').html(text);

    }

})();

function handleFileSelect(f) {

    var reader = new FileReader();
    reader.onload = function () {
        var text = reader.result;
        $('body').html(text);
		
		Office.context.document.settings.set('elementHTML', text);
        Office.context.document.settings.saveAsync();

    };
    reader.readAsText(f[0]);

}

function insURL() {

    var url = document.getElementById("web").value;

    $("body").html("Loading...");

    $.ajax({
        url: url,
        type: "GET",
        dataType: "html",
        headers: {
            'X-Requested-With': 'XMLHttpRequest'
        }
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

        Office.context.document.settings.set('elementHTML', $('body').html());
        Office.context.document.settings.saveAsync();

    }).fail(function (jqXHR, textStatus, errorThrown) {
        $("body").html("Error. File is not loaded");
    });

}