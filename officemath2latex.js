//
// officemath2latex.js (under construction)
//
// This program defines a function named processMathNode that converts mathematical
// expressions in OfficeMath format (DocumentFormat.OpenXml.Math) to LaTeX code.
// It takes an XML element as input and returns the corresponding LaTeX representation.
// The input mathXml element must have the node name "m:oMath".
//

function processMathNode(mathXml, chr = "") {
  console.log("processMathNode" + chr);
  let mathString = "";
  for (const mathElement of mathXml.childNodes) {
    console.log(mathElement);
    mathString += processMathElement(mathElement, chr);
    console.log("RESULT: " + mathString);
  }
  return mathString;
}

function processMathElement(mathElement, chr = "") {
  console.log("-processMathElement" + chr);
  switch (mathElement.nodeName) {
    case "m:e":
      return processMathNode(mathElement, chr) + chr;
    case "m:r":
      return processMath_run(mathElement, chr);
    case "m:sSup":
      return processMath_sSup(mathElement, chr);
    case "m:sSub":
      return processMath_sSub(mathElement, chr);
    case "m:d":
      return processMath_d(mathElement, chr);
    case "m:f":
      return processMath_f(mathElement, chr);
    case "m:nary":
      return processMath_nary(mathElement, chr);
    case "m:func":
      return processMath_func(mathElement, chr);
    case "m:limLow":
      return processMath_limLow(mathElement, chr);
    case "m:m":
      return processMath_m(mathElement, chr);
    case "m:acc":
      return processMath_acc(mathElement, chr);
    case "m:rad":
      return processMath_rad(mathElement, chr);
    default:
      break;
  }
  console.log("* not processed * " + mathElement.nodeName);
  return "";
}

function processMath_run(mathElements, chr = "") {
  console.log("--processMath_run" + chr);
  let mathString = "";
  let flagBold = false;

  for (const mathRunElement of mathElements.childNodes) {
    if (mathRunElement.nodeName === "m:rPr") {
      for (const rPr of mathRunElement.childNodes) {
        if (rPr.nodeName === "m:sty") {
          if (rPr.getAttribute("m:val") === "b") flagBold = true;
        }
      }
    } else if (mathRunElement.nodeName === "m:t") {
      if (mathRunElement.getAttribute("xml:space") === "preserve") {
        mathString += "\\ \\ ";
      } else {
        mathString += mathRunElement.textContent;
      }
    }
  }

  const replacestrs = [
    { pre: /π/g, post: "\\pi " },
    { pre: /∞/g, post: "\\infty " },
    { pre: /→/g, post: "\\rightarrow " },
    { pre: /±/g, post: "\\pm " },
    { pre: /∓/g, post: "\\mp " },
    { pre: /α/g, post: "\\alpha " },
    { pre: /β/g, post: "\\beta " },
    { pre: /…/g, post: "\\ldots " },
    { pre: /⋅/g, post: "\\cdot " },
    { pre: /×/g, post: "\\times " },
  ];

  for (const replacestr of replacestrs) {
    mathString = mathString.replace(replacestr.pre, replacestr.post);
  }
  if (flagBold) mathString = "\\mathbf{" + mathString + "}";

  return mathString;
}

function processMath_sSup(mathElements, chr = "") {
  console.log("-processMath_sSup" + chr);
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement, chr);
    } else if (mathElement.nodeName === "m:sup") {
      mathString += "^{" + processMathNode(mathElement, chr) + "}";
    }
  }
  return mathString;
}

function processMath_sSub(mathElements, chr = "") {
  console.log("-processMath_sSub" + chr);
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement, chr);
    } else if (mathElement.nodeName === "m:sub") {
      mathString += "_{" + processMathNode(mathElement, chr) + "}";
    }
  }
  return mathString;
}

function processMath_d(mathElements, chr = "") {
  console.log("-processMath_d" + chr);
  let mathString = "";
  let openPar = "\\left( ";
  let closePar = "\\right)";

  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:dPr") {
      for (const dPr of mathElement.childNodes) {
        if (dPr.nodeName === "m:begChr") {
          if (dPr.getAttribute("m:val") === "|") openPar = "\\left| ";
        } else if (dPr.nodeName === "m:endChr") {
          if (dPr.getAttribute("m:val") === "|") closePar = "\\right|";
        }
      }
    } else if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement, chr);
    }
  }
  if (mathString.startsWith("\\binom{")) return mathString;
  else return openPar + mathString + closePar;
}

function processMath_f(mathElements, chr = "") {
  console.log("-processMath_f" + chr);
  let mathString = "";
  let fracString = "\\frac{";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:fPr") {
      for (const fPr of mathElement.childNodes) {
        if (fPr.nodeName === "m:type") {
          if (fPr.getAttribute("m:val") === "noBar") fracString = "\\binom{";
        }
      }
    } else if (mathElement.nodeName === "m:num") {
      mathString += processMathNode(mathElement, chr);
      mathString += "}{";
    }
    if (mathElement.nodeName === "m:den") {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return fracString + mathString + "}";
}

function processMath_nary(mathElements, chr = "") {
  console.log("-processMath_nary" + chr);
  let mathString = "\\sum";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:sub") {
      mathString += "_{" + processMathNode(mathElement, chr);
      mathString += "}";
    } else if (mathElement.nodeName === "m:sup") {
      mathString += "^{" + processMathNode(mathElement, chr);
      mathString += "}";
    } else {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return mathString;
}

function processMath_func(mathElements, chr = "") {
  console.log("-processMath_func" + chr);
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:fName") {
      const functionName = processMathNode(mathElement, chr);
      switch (functionName) {
        case "cos":
          mathString += "\\cos ";
          break;
        case "sin":
          mathString += "\\sin ";
          break;
      }
    } else {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return mathString;
}

function processMath_limLow(mathElements, chr = "") {
  console.log("-processMath_limLow" + chr);
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      const elementString = processMathNode(mathElement, chr);
      if (elementString.trim() === "lim") mathString += "\\" + elementString.trim();
      else mathString += elementString;
    } else if (mathElement.nodeName === "m:lim") {
      mathString += "_{" + processMathNode(mathElement, chr) + "}";
    }
  }
  return mathString;
}

function processMath_m(mathElements, chr = "") {
  console.log("-processMath_m" + chr);
  let mathString = "\\begin{matrix}\n";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:mr") {
      mathString += processMathNode(mathElement, "&").slice(0, -1) + " \\\\ \n";
    }
  }
  return mathString + "\\end{matrix} ";
}

function processMath_acc(mathElements, chr = "") {
  console.log("-processMath_acc" + chr);
  let mathString = "\\overrightarrow{";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return mathString + "}";
}

function processMath_rad(mathElements, chr = "") {
  console.log("-processMath_rad" + chr);
  let mathString = "\\sqrt{";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return mathString + "}";
}
