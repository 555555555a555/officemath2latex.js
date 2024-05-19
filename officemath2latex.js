//
// officemath2lateX.js (under construction)
//
// This program defines a function named processMathLoop(mathXml) that converts mathematical
// expressions in OfficeMath format (likely from DocumentFormat.OpenXml.Math) to LaTeX code.
// It takes an XML element as input and returns the corresponding LaTeX representation.
// The input mathXml element must have the node name "m:oMath".
//

function processMathLoop(mathXml) {
  console.log("processMathLoop");
  let mathString = "";
  for (const mathElement of mathXml.childNodes) {
    console.log(mathElement);
    mathString += processMathElement(mathElement);
    console.log("RESULT: " + mathString);
  }
  return mathString;
}

function processMathElement(mathElement) {
  console.log("-processMathElement");
  switch (mathElement.nodeName) {
    case "m:e":
      return processMathLoop(mathElement);
    case "m:r":
      return processMath_run(mathElement);
    case "m:sSup":
      return processMath_sSup(mathElement);
    case "m:sSub":
      return processMath_sSub(mathElement);
    case "m:d":
      return processMath_d(mathElement);
    case "m:f":
      return processMath_f(mathElement);
    case "m:nary":
      return processMath_nary(mathElement);
    case "m:func":
      return processMath_func(mathElement);
    case "m:limLow":
      return processMath_limLow(mathElement);
    default:
      break;
  }
  console.log("* not processed * " + mathElement.nodeName);
  return "";
}

function processMath_run(mathElements) {
  console.log("--processMath_run");
  let mathString = "";
  for (const mathRunElement of mathElements.childNodes) {
    if (mathRunElement.nodeName === "m:t") {
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
  ];

  for (const replacestr of replacestrs) {
    mathString = mathString.replace(replacestr.pre, replacestr.post);
  }
  return mathString;
}

function processMath_sSup(mathElements) {
  console.log("-processMath_sSup");
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathLoop(mathElement);
    } else if (mathElement.nodeName === "m:sup") {
      mathString += "^{" + processMathLoop(mathElement) + "}";
    }
  }
  return mathString;
}

function processMath_sSub(mathElements) {
  console.log("-processMath_sSub");
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathLoop(mathElement);
    } else if (mathElement.nodeName === "m:sub") {
      mathString += "_{" + processMathLoop(mathElement) + "}";
    }
  }
  return mathString;
}

function processMath_d(mathElements) {
  console.log("-processMath_d");
  let mathString = "\\left( ";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathLoop(mathElement);
    }
  }
  return mathString + "\\right)";
}

function processMath_f(mathElements) {
  console.log("-processMath_f");
  let mathString = "\\frac{";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:num") {
      mathString += processMathLoop(mathElement);
      mathString += "}{";
    }
    if (mathElement.nodeName === "m:den") {
      mathString += processMathLoop(mathElement) + "}";
    }
  }
  return mathString;
}

function processMath_nary(mathElements) {
  console.log("-processMath_nary");
  let mathString = "\\sum";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:sub") {
      mathString += "_{" + processMathLoop(mathElement);
      mathString += "}";
    } else if (mathElement.nodeName === "m:sup") {
      mathString += "^{" + processMathLoop(mathElement);
      mathString += "}";
    } else {
      mathString += processMathLoop(mathElement);
    }
  }
  return mathString;
}

function processMath_func(mathElements) {
  console.log("-processMath_func");
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:fName") {
      const functionName = processMathLoop(mathElement);
      switch (functionName) {
        case "cos":
          mathString += "\\cos ";
          break;
        case "sin":
          mathString += "\\sin ";
          break;
      }
    } else {
      mathString += processMathLoop(mathElement);
    }
  }
  return mathString;
}

function processMath_limLow(mathElements) {
  console.log("-processMath_limLow");
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      const elementString = processMathLoop(mathElement);
      if (elementString.trim() === "lim") mathString += "\\" + elementString.trim();
      else mathString += elementString;
    } else if (mathElement.nodeName === "m:lim") {
      mathString += "_{" + processMathLoop(mathElement) + "}";
    }
  }
  return mathString;
}
