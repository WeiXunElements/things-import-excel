# things-import-excel

## Import Excel은 JS-XLSX를 이용하여 정형화된 Excel을 json 또는 CSV로 바꾸는 도구다.

    해당 컴포넌트을 사용하고 import Function을 호출해 줘야 Import 화면이 나온다.

    Excel 파싱이 완료되면 `excel-process-finish` Event가 발생된다.

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

Element dependencies are managed via [Bower](http://bower.io/). You can
install that via:

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
