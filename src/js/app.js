import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import "materialize-css/dist/css/materialize.min.css";
import "materialize-css/dist/js/materialize.min";
import { merge } from "lodash";

var expressions = require("angular-expressions");

// The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
function replaceErrors(value) {
  if (value instanceof Error) {
    return Object.getOwnPropertyNames(value).reduce(function (error, key) {
      error[key] = value[key];
      return error;
    }, {});
  }
  return value;
}

expressions.filters.lower = function (input) {
  // This condition should be used to make sure that if your input is
  // undefined, your output will be undefined as well and will not
  // throw an error
  if (!input) return input;
  return input.toLowerCase();
};

function angularParser(tag) {
  if (tag === ".") {
    return {
      get: function (s) {
        return s;
      },
    };
  }
  const expr = expressions.compile(
    tag.replace(/(’|‘)/g, "'").replace(/(“|”)/g, '"')
  );
  return {
    get: function (scope, context) {
      let obj = {};
      const scopeList = context.scopeList;
      const num = context.num;
      for (let i = 0, len = num + 1; i < len; i++) {
        obj = merge(obj, scopeList[i]);
      }
      return expr(scope, obj);
    },
  };
}

var output = {};
(function () {
  function toJSONString(form) {
    var obj = {};
    var elements = form.querySelectorAll("input, select, textarea");
    for (var i = 0; i < elements.length; ++i) {
      var element = elements[i];
      var name = element.name;
      var value = String(element.value);
      if (element.type && element.type === "checkbox") {
        if (element.checked) {
          value = true;
        } else {
          value = false;
        }
      }

      if (name) {
        obj[name] = value;
      }
    }

    return JSON.stringify(obj);
  }

  document.addEventListener("DOMContentLoaded", function () {
    var form = document.getElementById("employmentForm");
    form.addEventListener(
      "submit",
      function (e) {
        e.preventDefault();
        var json = toJSONString(this);
        output = json;
        generate();
      },
      false
    );
  });
})();

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}
function generate() {
  loadFile("contract_templates/Employment Agreement.docx", function (
    error,
    content
  ) {
    if (error) {
      throw error;
    }

    // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
    function replaceErrors(value) {
      if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function (error, key) {
          error[key] = value[key];
          return error;
        }, {});
      }
      return value;
    }

    function errorHandler(error) {
      console.log(JSON.stringify({ error: error }, replaceErrors));

      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map(function (error) {
            return error.properties.explanation;
          })
          .join("\n");
        console.log("errorMessages", errorMessages);
        // errorMessages is a humanly readable message looking like this :
        // 'The tag beginning with "foobar" is unopened'
      }
      throw error;
    }

    var zip = new PizZip(content);
    var doc;
    try {
      doc = new Docxtemplater(zip, { parser: angularParser });
    } catch (error) {
      // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
      errorHandler(error);
    }

    doc.setData(JSON.parse(output));
    try {
      doc.render();
    } catch (error) {
      // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
      errorHandler(error);
    }

    var out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }); //Output the document using Data-URI
    var JSONoutput = JSON.parse(output);
    console.log(JSONoutput.employeeName);
    var documentTitle =
      JSONoutput.employeeName +
      " - " +
      JSONoutput.employerName +
      " Employment Agreement.docx";
    saveAs(out, documentTitle);
  });
}
