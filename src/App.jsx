import { useState, useEffect } from "react";
import { saveAs } from "file-saver";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const DocxEditor = () => {
  const [formData, setFormData] = useState({
    CustomerName: "",
    ReferenceNo: "",
    Date: "",
    CustomerPhone: "",
    CustomerAddress: "",
    CompanyPOC: "",
    CompanyName: "",
    CompanyPhone: "",
    CompanyAddress: "",
    ProjectSize: "",
  });

  const [docFile, setDocFile] = useState(null);

  useEffect(() => {
    const loadTemplate = async () => {
      const response = await fetch("/demo.docx"); 
      const blob = await response.blob();
      const reader = new FileReader();

      reader.onload = function (event) {
        const content = event.target.result;
        const zip = new PizZip(content);
        try {
          const doc = new Docxtemplater(zip, {
            delimiters: { start: "<<", end: ">>" },
          });
          setDocFile(doc);
          console.log("Document loaded successfully.");
        } catch (error) {
          console.error("Error initializing Docxtemplater:", error);
          alert(
            "Failed to load the document. Please ensure it is a valid .docx template."
          );
        }
      };

      reader.readAsBinaryString(blob);
    };

    loadTemplate();
  }, []);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData((prevData) => ({
      ...prevData,
      [name]: value,
    }));
  };

  const handleGenerateDoc = () => {
    if (!docFile) {
      alert("Template is loading, please wait...");
      return;
    }
    docFile.setData({
      CustomerName: formData.CustomerName || "Default Customer Name",
      ReferenceNo: formData.ReferenceNo || "Default Reference",
      Date: formData.Date || "Default Date",
      CustomerPhone: formData.CustomerPhone || "Default Phone",
      CustomerAddress: formData.CustomerAddress || "Default Address",
      CompanyPOC: formData.CompanyPOC || "Default POC",
      CompanyName: formData.CompanyName || "Default Company",
      CompanyPhone: formData.CompanyPhone || "Default Company Phone",
      CompanyAddress: formData.CompanyAddress || "Default Company Address",
      ProjectSize: formData.ProjectSize || "Default Project Size",
    });

    try {
      docFile.render();
      console.log("Document rendered successfully.");
    } catch (error) {
      if (error.properties && error.properties.errors) {
        console.error("Template rendering errors:", error.properties.errors);
        alert(
          "There were errors in the template:\n" +
            error.properties.errors.map((err) => `- ${err}`).join("\n")
        );
      } else {
        console.error("Error rendering docx:", error);
        alert(
          "An error occurred while rendering the document. Check the console for details."
        );
      }
      return;
    }
    const output = docFile.getZip().generate({ type: "blob" });
    saveAs(output, "generated-proposal.docx");
  };

  return (
    <div className="max-w-2xl mx-auto p-6 bg-white shadow-md rounded-lg mt-10">
      <h1 className="text-2xl font-bold mb-4 text-center">Proposals</h1>
      <div>
        <h2 className="text-lg font-semibold mb-2">Client Information</h2>
        <div className="grid grid-cols-1 sm:grid-cols-3 md:grid-cols-2 lg:grid-cols-2 gap-5">
          <input
            type="text"
            name="CustomerName"
            placeholder="Enter Customer Name"
            value={formData.CustomerName}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="ReferenceNo"
            placeholder="Enter Reference No"
            value={formData.ReferenceNo}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="date"
            name="Date"
            placeholder="Select Date"
            value={formData.Date}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="CustomerPhone"
            placeholder="Enter Customer Phone"
            value={formData.CustomerPhone}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="CustomerAddress"
            placeholder="Enter Customer Address"
            value={formData.CustomerAddress}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="ProjectSize"
            placeholder="Enter Project Size (kW)"
            value={formData.ProjectSize}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
        </div>

        <h2 className="text-lg font-semibold mb-2 mt-6">Company Information</h2>
        <div className="grid grid-cols-1 sm:grid-cols-3 md:grid-cols-2 lg:grid-cols-2 gap-5">
          <input
            type="text"
            name="CompanyPOC"
            placeholder="Enter Company POC"
            value={formData.CompanyPOC}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="CompanyName"
            placeholder="Enter Company Name"
            value={formData.CompanyName}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="CompanyPhone"
            placeholder="Enter Company Phone"
            value={formData.CompanyPhone}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
          <input
            type="text"
            name="CompanyAddress"
            placeholder="Enter Company Address"
            value={formData.CompanyAddress}
            onChange={handleChange}
            className="p-2 border border-gray-300 rounded-md focus:outline-none focus:ring focus:ring-orange-300"
          />
        </div>
      </div>
      <button
        onClick={handleGenerateDoc}
        className="mt-4 w-full py-2 bg-orange-400 text-white font-bold rounded-md hover:bg-orange-500 transition duration-300"
      >
        Generate DOCX
      </button>
    </div>
  );
};

export default DocxEditor;
