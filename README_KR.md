# things-import-excel

## Import Excel은 JS-XLSX를 이용하여 정형화된 Excel을 json 또는 CSV로 바꾸는 도구다.

    해당 컴포넌트를 사용하고 import Function을 호출해 줘야 Import 화면이 나온다.

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

element의 종속성은 [Bower](http://bower.io/)를 통해 관리되며, 아래의 방법을 통해 설치할 수 있다.

    npm install -g bower

그리고, 아래의 방법을 통해 실행할 수 있다.

    bower install

## Playing With Your Element

element를 독립적으로 처리하려면 [Polyserve](https://github.com/PolymerLabs/polyserve)를 사용하여 element의 bower 의존성을 유지하도록 하며, 이는 아래의 방법을 통해 설치할 수 있다.

    npm install -g polymer-cli

그리고, 아래의 방법을 통해 실행할 수 있다.

    polymer serve

element를 실행한 경우, `things-import-excel`이 디렉토리 이름으로 포함되어 있는 `http://localhost:8080/components/things-import-excel/`을 통해 이를 미리 확인할 수 있다. 
