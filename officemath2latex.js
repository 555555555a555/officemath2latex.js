//
// officemath2latex.js
//
// This program defines a function named processMathNode that converts mathematical
// expressions in OfficeMath format (DocumentFormat.OpenXml.Math) to LaTeX code.
// It takes an XML element as input and returns the corresponding LaTeX representation.
// The input mathXml element must be "m:oMath" node.
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
    case "m:eqArr":
      return processMath_eqArr(mathElement, chr);
    case "m:box":
      return "{" + processMathNode(mathElement, chr) + "}";
    case "m:sPre":
      return processMath_sPre(mathElement, chr);
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

  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:rPr") {
      for (const rPr of mathElement.childNodes) {
        if (rPr.nodeName === "m:sty") {
          if (rPr.getAttribute("m:val") === "b") flagBold = true;
        }
      }
    } else if (mathElement.nodeName === "m:t") {
      if (mathElement.getAttribute("xml:space") === "preserve") {
        if (mathElement.textContent.trim() === "") mathString += "\\ \\ ";
        else mathString += processMath_FieldCode(mathElement.textContent);
      } else {
        mathString += mathElement.textContent;
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
    { pre: /γ/g, post: "\\gamma " },
    { pre: /…/g, post: "\\ldots " },
    { pre: /⋅/g, post: "\\cdot " },
    { pre: /×/g, post: "\\times " },
    { pre: /θ/g, post: "\\theta " },
    { pre: /Γ/g, post: "\\Gamma " },
    { pre: /≈/g, post: "\\approx " },
    { pre: /ⅈ/g, post: "i " }, //"\\mathbb{i} "? mathbb -> usepackage{amsfonts}
    { pre: /∇/g, post: "\\nabla " },
    { pre: /ⅆ/g, post: "d " }, //"\\mathbb{d} "?
    { pre: /≥/g, post: "\\geq " },
    { pre: /∀/g, post: "\\forall " },
    { pre: /∃/g, post: "\\exists " },
    { pre: /∧/g, post: "\\land " },
    { pre: /⇒/g, post: "\\Rightarrow " },
    { pre: /ψ/g, post: "\\psi " },
    { pre: /∂/g, post: "\\partial " },
    { pre: /≠/g, post: "\\neq " },
    { pre: /~/g, post: "\\sim " },
    { pre: /÷/g, post: "\\div " },
    { pre: /∝/g, post: "\\propto " },
    { pre: /≪/g, post: "\\ll " },
    { pre: /≫/g, post: "\\gg " },
    { pre: /≤/g, post: "\\leq " },
    { pre: /≅/g, post: "\\cong " },
    { pre: /≡/g, post: "\\equiv " },
    { pre: /∁/g, post: "\\complement " },
    { pre: /∪/g, post: "\\cup " },
    { pre: /∩/g, post: "\\cap " },
    { pre: /∅/g, post: "\\varnothing " },
    { pre: /∆/g, post: "\\mathrm{\\Delta}, " },
    { pre: /∄/g, post: "\\nexists " },
    { pre: /∈/g, post: "\\in " },
    { pre: /∋/g, post: "\\ni " },
    { pre: /←/g, post: "\\leftarrow " },
    { pre: /↑/g, post: "\\uparrow " },
    { pre: /↓/g, post: "\\downarrow " },
    { pre: /↔/g, post: "\\leftrightarrow " },
    { pre: /∴/g, post: "\\therefore " },
    { pre: /¬/g, post: "\\neg " },
    { pre: /δ/g, post: "\\delta " },
    { pre: /ε/g, post: "\\varepsilon " },
    { pre: /ϵ/g, post: "\\epsilon " },
    { pre: /ϑ/g, post: "\\vartheta " },
    { pre: /μ/g, post: "\\mu " },
    { pre: /ρ/g, post: "\\rho " },
    { pre: /σ/g, post: "\\sigma " },
    { pre: /τ/g, post: "\\tau " },
    { pre: /φ/g, post: "\\varphi " },
    { pre: /ω/g, post: "\\omega " },
    { pre: /∙/g, post: "\\bullet " },
    { pre: /⋮/g, post: "\\vdots " },
    { pre: /⋱/g, post: "\\ddots " },
    { pre: /ℵ/g, post: "\\aleph " },
    { pre: /ℶ/g, post: "\\beth " },
    { pre: /∎/g, post: "\\blacksquare " },
    { pre: /%°/g, post: "\\%{^\\circ} " },
    { pre: /√/g, post: "\\sqrt{} " },
    { pre: /∛/g, post: "\\sqrt[3]{} " },
    { pre: /∜/g, post: "\\sqrt[4]{} " },
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
          if (dPr.getAttribute("m:val") === "{") openPar = "\\left\\{ ";
          if (dPr.getAttribute("m:val") === "[") openPar = "\\left\\lbrack ";
          if (dPr.getAttribute("m:val") === "") openPar = "\\left. ";
        } else if (dPr.nodeName === "m:endChr") {
          if (dPr.getAttribute("m:val") === "|") closePar = "\\right|";
          if (dPr.getAttribute("m:val") === "}") closePar = "\\right\\} ";
          if (dPr.getAttribute("m:val") === "]") closePar = "\\right\\rbrack ";
          if (dPr.getAttribute("m:val") === "") closePar = "\\right. ";
        }
      }
    } else if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement, chr);
    }
  }
  if (mathString.startsWith("\\genfrac{}{}{0pt}{}{")) return mathString.replace("\\genfrac{}{}{0pt}{}{", "\\binom{");
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
          if (fPr.getAttribute("m:val") === "noBar") fracString = "\\genfrac{}{}{0pt}{}{"; // if in d "\\binom{";
          if (fPr.getAttribute("m:val") === "skw") fracString = "\\nicefrac{"; //\usepackage{units}
          if (fPr.getAttribute("m:val") === "lin") fracString = " ";
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
  if (fracString === " ") return "\\ " + mathString.replace("}{", "/");
  return fracString + mathString + "}";
}

function processMath_nary(mathElements, chr = "") {
  console.log("-processMath_nary" + chr);
  let mathString = "\\int";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:naryPr") {
      for (const naryPr of mathElement.childNodes) {
        if (naryPr.nodeName === "m:chr") {
          if (naryPr.getAttribute("m:val") === "∑") mathString = "\\sum";
          else if (naryPr.getAttribute("m:val") === "∏") mathString = "\\prod";
          else if (naryPr.getAttribute("m:val") === "∐") mathString = "\\coprod";
          else if (naryPr.getAttribute("m:val") === "∬") mathString = "\\iint";
          else if (naryPr.getAttribute("m:val") === "∭") mathString = "\\iiint";
          else if (naryPr.getAttribute("m:val") === "∮") mathString = "\\oint";
          else if (naryPr.getAttribute("m:val") === "∯") mathString = "\\oiint"; //\usepackage{esint}
          else if (naryPr.getAttribute("m:val") === "∰") mathString = "\\oiiint"; // \usepackage{mathdesign,mdsymbol}
          else if (naryPr.getAttribute("m:val") === "⋃") mathString = "\\bigcup";
          else if (naryPr.getAttribute("m:val") === "⋂") mathString = "\\bigcap";
          else if (naryPr.getAttribute("m:val") === "⋁") mathString = "\\bigvee";
          else if (naryPr.getAttribute("m:val") === "⋀") mathString = "\\bigwedge";
        }
      }
    } else if (mathElement.nodeName === "m:sub") {
      mathString += "_{" + processMathNode(mathElement, chr);
      mathString += "}";
    } else if (mathElement.nodeName === "m:sup") {
      mathString += "^{" + processMathNode(mathElement, chr);
      mathString += "}";
    } else {
      mathString += "{" + processMathNode(mathElement, chr) + "}";
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
          mathString += "\\cos{ ";
          break;
        case "sin":
          mathString += "\\sin{ ";
          break;
        default:
          mathString += functionName + "{";
          break;
      }
    } else {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return mathString + "}";
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
  let mathString = "\\sqrt";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:deg") {
      mathString += "[" + processMathNode(mathElement, chr) + "]";
    }
    if (mathElement.nodeName === "m:e") {
      mathString += "{" + processMathNode(mathElement, chr) + "}";
    }
  }
  return mathString;
}

function processMath_eqArr(mathElements, chr = "") {
  console.log("-processMath_eqArr" + chr);
  let mathString = "\\begin{aligned}\n";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:e") {
      mathString += processMathNode(mathElement) + " \\\\ \n";
    }
  }
  mathString = mathString.replace(/ &/g, "\\ \\ &");
  if (mathString.endsWith(" \\\\ \n")) mathString = mathString.slice(0, -5) + "\n";
  return mathString + "\\end{aligned} ";
}

function processMath_sPre(mathElements, chr = "") {
  console.log("-processMath_sPre" + chr);
  let mathString = "";
  for (const mathElement of mathElements.childNodes) {
    if (mathElement.nodeName === "m:sub") {
      mathString += "_{" + processMathNode(mathElement, chr);
      mathString += "}";
    } else if (mathElement.nodeName === "m:sup") {
      mathString += "^{" + processMathNode(mathElement, chr);
      mathString += "}";
    } else if (mathElement.nodeName !== "m:ctrlPr") {
      mathString += processMathNode(mathElement, chr);
    }
  }
  return mathString;
}

function processMath_FieldCode(mathText, chr = "") {
  console.log("-processMath_FieldCode" + chr);

  const results_cancel1 = mathText.match(/eq \\o\s?\((.*?),\/\)/i);
  const results_cancel2 = mathText.match(/eq \\o\s?\((.*?),／\)/i);
  const results_overline = mathText.match(/eq \\x\s?\\to \((.*?)\)/i);

  if (results_cancel1) {
    return `\\cancel{${results_cancel1[1]}}`;
  }
  if (results_cancel2) {
    return `\\cancel{${results_cancel2[1]}}`;
  }
  if (results_overline) {
    return `\\overline{${results_overline[1]}}`;
  }

  return mathText;
}
