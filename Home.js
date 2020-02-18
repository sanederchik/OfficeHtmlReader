'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            try {

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
                    '               <p> More on my <a href="https://github.com/sanederchik/OfficeHtmlReader"> github page </a> </p>  ' +
                    '           </div>  ' +
                    '   		  ' +
                    '      </div>  '; 
                var t1 = Office.context.document.settings.get('elementHTML');
           


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

    	var _ = document.getElementById('web');
    	$("body").html("Loading...");

	if (_.value.toString() == ''){

	alert('Пустое значение!');
	} else {

	document.body.innerHTML = `

	<iframe src = "${_.value}" width="100%" height="100%" frameBorder="0"><Браузер не поддерживает iframe</iframe>

	`;
	}

}
