# things-import-excel

## Import Excel is a tool that converts standardized Excel into json or CSV using JS-XLSX.

    The Import screen will appear only if you use the corresponding component and call the import function.

    When Excel parsing is completed, an `excel-process-finish` event is generated.

Example:

```html
    <paper-input id="result" label="Result" type="text"></paper-input>
    <things-import-excel id="import-excel" handle-as="json" parsed-result="{{result}}"> </things-import-excel>
    <paper-button raised id="import-btn"> importXLSX </paper-button>
    <script>
      var importBtn = document.querySelector("#import-btn");
      importBtn.addEventListener('tap', importXLSX);

      var importExcel = document.querySelector("#import-excel");
      importExcel.addEventListener('excel-process-finish', showResult)

      var result = document.querySelector("#result")

      function importXLSX(e) {
          importExcel.import();
      };

      function showResult (e) {
         result.value = JSON.stringify(demo.result);
      }

    </script>

```

## Dependencies

Element dependencies are managed via [Bower](http://bower.io/). You can install that via:

    npm install -g bower

Then, go ahead and download the element's dependencies:

    bower install

## Playing With Your Element

If you wish to work on your element in isolation, we recommend that you use
[Polyserve](https://github.com/PolymerLabs/polyserve) to keep your element's
bower dependencies in line. You can install it via:

    npm install -g polymer-cli

And you can run it via:

    polymer serve

Once running, you can preview your element at
`http://localhost:8080/components/things-import-excel/`, where `things-import-excel` is the name of the directory containing it.
