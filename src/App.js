import "./styles.css";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";
import Template from "./assets/officialClaim.docx";

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

export default function App() {
  const date = new Date();

  const generateDocument = () => {
    loadFile(Template, function (error, content) {
      if (error) {
        throw error;
      }
      var zip = new PizZip(content);
      var doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });
      doc.setData({
        signing_location: "Toulouse",
        signing_date: `${date.getDate()}/${date.getMonth()}/${date.getFullYear()}`,
        first_name: "Sébastien",
        last_name: "François",
        birth_date: "19 août 1994",
        claim_type: "Casse",
        incident_location: "3 rue Ella Maillart, 31300 Toulouse",
        incident_timestamp: new Date().getFullYear(),
        bike_name: "Vélo de champion",
        incident_description:
          "Faits et actions avant, pendant, après le sinistre",
        customer_name: "Morio"
      });
      try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render();
      } catch (error) {
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
        function replaceErrors(key, value) {
          if (value instanceof Error) {
            return Object.getOwnPropertyNames(value).reduce(function (
              error,
              key
            ) {
              error[key] = value[key];
              return error;
            },
            {});
          }
          return value;
        }
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
      var out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      }); //Output the document using Data-URI
      saveAs(out, "déclaration_circonstanciée.docx");
    });
  };
  return (
    <div className="p-2">
      <button onClick={generateDocument}>Generate document</button>
    </div>
  );
}
