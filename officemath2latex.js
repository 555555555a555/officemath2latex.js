//
// officemath2latex.js
//
// This program defines a function named processMathNode that converts mathematical
// expressions in OfficeMath format (DocumentFormat.OpenXml.Math) to LaTeX code.
// It takes an XML element as input and returns the corresponding LaTeX representation.
// The input mathXml element must be "m:oMath" node.
//

function processMathNode(mathXml, chr = "") {
  let mathString = "";
  for (const childXml of mathXml.childNodes) {
    const childElement = new OfficeMathElement(childXml);
    mathString += childElement.process(chr);
  }
  return mathString;
}

class OfficeMathElement {
  constructor(node) {
    this.node = node;
  }

  process(chr = "") {
    switch (this.node.nodeName) {
      case "m:e":
        return processMathNode(this.node, chr) + chr;
      case "m:r":
        return new OfficeMathRun(this.node).process(chr);
      case "m:sSup":
        return new OfficeMathSuperscriptSubscript(this.node).process(chr, "^");
      case "m:sSub":
        return new OfficeMathSuperscriptSubscript(this.node).process(chr, "_");
      case "m:d":
        return new OfficeMathDelimiter(this.node).process(chr);
      case "m:f":
        return new OfficeMathFraction(this.node).process(chr);
      case "m:nary":
        return new OfficeMathNary(this.node).process(chr);
      case "m:func":
        return new OfficeMathFunction(this.node).process(chr);
      case "m:limLow":
        return new OfficeMathLimLow(this.node).process(chr);
      case "m:m":
        return new OfficeMathMatrix(this.node).process(chr);
      case "m:acc":
        return new OfficeMathAccent(this.node).process(chr);
      case "m:rad":
        return new OfficeMathRadical(this.node).process(chr);
      case "m:eqArr":
        return new OfficeMathEquationArray(this.node).process(chr);
      case "m:box":
        return `{${processMathNode(this.node, chr)}}`;
      case "m:sPre":
        return new OfficeMathSPre(this.node).process(chr);
      case "m:t":
        return new OfficeMathText(this.node).process(chr);
      default:
        return "";
    }
  }
}

class OfficeMathSuperscriptSubscript extends OfficeMathElement {
  process(chr, type) {
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:e") {
        mathString += processMathNode(mathElement, chr);
      } else if (mathElement.nodeName === "m:sup" && type === "^") {
        mathString += `^{${processMathNode(mathElement, chr)}}`;
      } else if (mathElement.nodeName === "m:sub" && type === "_") {
        mathString += `_{${processMathNode(mathElement, chr)}}`;
      }
    }
    return mathString;
  }
}

class OfficeMathDelimiter extends OfficeMathElement {
  process(chr) {
    let mathString = "";
    let openPar = "\\left( ";
    let closePar = "\\right)";

    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:dPr") {
        const delimiters = this.getDelimiters(mathElement);
        openPar = delimiters.openPar;
        closePar = delimiters.closePar;
      } else if (mathElement.nodeName === "m:e") {
        mathString += processMathNode(mathElement, chr);
      }
    }

    if (mathString.startsWith("\\genfrac{}{}{0pt}{}")) {
      return mathString.replace("\\genfrac{}{}{0pt}{}", "\\binom");
    }

    return `${openPar}${mathString}${closePar}`;
  }

  getDelimiters(dPrNode) {
    let openPar = "\\left( ";
    let closePar = "\\right)";

    for (const dPr of dPrNode.childNodes) {
      if (dPr.nodeName === "m:begChr") {
        const val = dPr.getAttribute("m:val");
        openPar = val === "|" ? "\\left| " : val === "{" ? "\\left\\{ " : val === "[" ? "\\left\\lbrack " : "\\left. ";
      } else if (dPr.nodeName === "m:endChr") {
        const val = dPr.getAttribute("m:val");
        closePar =
          val === "|" ? "\\right|" : val === "}" ? "\\right\\} " : val === "]" ? "\\right\\rbrack " : "\\right. ";
      }
    }

    return { openPar, closePar };
  }
}

class OfficeMathRun extends OfficeMathElement {
  process(chr = "") {
    let mathString = "";
    let flagBold = false;

    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:rPr") {
        for (const rPr of mathElement.childNodes) {
          if (rPr.nodeName === "m:sty" && rPr.getAttribute("m:val") === "b") {
            flagBold = true;
          }
        }
      } else if (mathElement.nodeName === "m:t") {
        const textContent = mathElement.textContent.trim();
        if (mathElement.getAttribute("xml:space") === "preserve") {
          mathString += textContent === "" ? "\\ \\ " : new OfficeMathFieldCodeText(textContent).process(chr);
        } else {
          mathString += textContent;
        }
      }
    }

    const replacements = [
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
      { pre: /ⅈ/g, post: "i " }, // "\\mathbb{i} "?, in that case add \usepackage{amsfonts}
      { pre: /∇/g, post: "\\nabla " },
      { pre: /ⅆ/g, post: "d " }, // "\\mathbb{i} "?, in that case add \usepackage{amsfonts}
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

    const replacer = (str) => replacements.reduce((acc, { pre, post }) => acc.replace(pre, post), str);
    mathString = replacer(mathString);

    if (flagBold) {
      mathString = `\\mathbf{${mathString}}`;
    }

    return mathString;
  }
}

class OfficeMathFraction extends OfficeMathElement {
  process(chr) {
    const fracType = this.getFractionType();
    return this.processFraction(chr, fracType);
  }

  getFractionType() {
    let fracType = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:fPr") {
        for (const fPr of mathElement.childNodes) {
          if (fPr.nodeName === "m:type") {
            fracType = fPr.getAttribute("m:val");
          }
        }
      }
    }
    return fracType;
  }

  processFraction(chr, fracType) {
    let mathString = "";
    let numString = "";
    let denString = "";

    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:num") {
        numString = processMathNode(mathElement, chr);
      } else if (mathElement.nodeName === "m:den") {
        denString = processMathNode(mathElement, chr);
      }
    }

    switch (fracType) {
      case "noBar":
        // if this fraction is in delimiter, it will be replaced to \\binom in MathDelimiter.process
        mathString = `\\genfrac{}{}{0pt}{}{${numString}}{${denString}}`;
        break;
      case "skw":
        // to use \\nicefrac, add \usepackage{units}
        mathString = `\\nicefrac{${numString}}{${denString}}`;
        break;
      case "lin":
        mathString = `${numString}/${denString}`;
        break;
      default:
        mathString = `\\frac{${numString}}{${denString}}`;
    }

    return mathString;
  }
}

class OfficeMathNary extends OfficeMathElement {
  process(chr) {
    let mathString = "\\int";
    let subString = "";
    let supString = "";
    let postString = "";

    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:naryPr") {
        for (const naryPr of mathElement.childNodes) {
          if (naryPr.nodeName === "m:chr") {
            const val = naryPr.getAttribute("m:val");
            mathString = this.getNaryOperator(val);
          }
        }
      } else if (mathElement.nodeName === "m:sub") {
        subString = `_{${processMathNode(mathElement, chr)}}`;
      } else if (mathElement.nodeName === "m:sup") {
        supString = `^{${processMathNode(mathElement, chr)}}`;
      } else {
        postString += `{${processMathNode(mathElement, chr)}}`;
      }
    }

    return `${mathString}${subString}${supString}${postString}`;
  }

  getNaryOperator(val) {
    switch (val) {
      case "∑":
        return "\\sum";
      case "∏":
        return "\\prod";
      case "∐":
        return "\\coprod";
      case "∬":
        return "\\iint";
      case "∭":
        return "\\iiint";
      case "∮":
        return "\\oint";
      case "∯":
        return "\\oiint"; // add \usepackage{esint}
      case "∰":
        return "\\oiiint"; // add \usepackage{mathdesign,mdsymbol}
      case "⋃":
        return "\\bigcup";
      case "⋂":
        return "\\bigcap";
      case "⋁":
        return "\\bigvee";
      case "⋀":
        return "\\bigwedge";
      default:
        return "\\int";
    }
  }
}

class OfficeMathFunction extends OfficeMathElement {
  process(chr) {
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:fName") {
        const functionName = processMathNode(mathElement, chr);
        mathString += this.getFunctionString(functionName);
      } else {
        mathString += processMathNode(mathElement, chr);
      }
    }
    return `${mathString}}`;
  }

  getFunctionString(functionName) {
    switch (functionName) {
      case "cos":
        return "\\cos{";
      case "sin":
        return "\\sin{";
      default:
        return `${functionName}{`;
    }
  }
}

class OfficeMathLimLow extends OfficeMathElement {
  process(chr) {
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:e") {
        const elementString = processMathNode(mathElement, chr);
        mathString += elementString.trim() === "lim" ? `\\${elementString.trim()}` : elementString;
      } else if (mathElement.nodeName === "m:lim") {
        mathString += `_{${processMathNode(mathElement, chr)}}`;
      }
    }
    return mathString;
  }
}

class OfficeMathMatrix extends OfficeMathElement {
  process(chr) {
    let mathString = "\\begin{matrix}\n";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:mr") {
        mathString += `${processMathNode(mathElement, "&").slice(0, -1)} \\\\ \n`;
      }
    }
    return `${mathString}\\end{matrix} `;
  }
}

class OfficeMathAccent extends OfficeMathElement {
  process(chr) {
    let mathString = "\\overrightarrow{";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:e") {
        mathString += processMathNode(mathElement, chr);
      }
    }
    return `${mathString}}`;
  }
}

class OfficeMathRadical extends OfficeMathElement {
  process(chr) {
    let mathString = "\\sqrt";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:deg") {
        mathString += `[${processMathNode(mathElement, chr)}]`;
      } else if (mathElement.nodeName === "m:e") {
        mathString += `{${processMathNode(mathElement, chr)}}`;
      }
    }
    return mathString;
  }
}

class OfficeMathEquationArray extends OfficeMathElement {
  process(chr) {
    let mathString = "\\begin{aligned}\n";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:e") {
        mathString += `${processMathNode(mathElement)} \\\\ \n`;
      }
    }
    mathString = mathString.replace(/ &/g, "\\ \\ &");
    if (mathString.endsWith(" \\\\ \n")) {
      mathString = mathString.slice(0, -5) + "\n";
    }
    return `${mathString}\\end{aligned} `;
  }
}

class OfficeMathSPre extends OfficeMathElement {
  process(chr) {
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:sub") {
        mathString += `_{${processMathNode(mathElement, chr)}}`;
      } else if (mathElement.nodeName === "m:sup") {
        mathString += `^{${processMathNode(mathElement, chr)}}`;
      } else if (mathElement.nodeName !== "m:ctrlPr") {
        mathString += processMathNode(mathElement, chr);
      }
    }
    return mathString;
  }
}

class OfficeMathText extends OfficeMathElement {
  process(chr) {
    const textContent = this.node.textContent.trim();
    if (this.node.getAttribute("xml:space") === "preserve") {
      return textContent === "" ? "\\ \\ " : new OfficeMathFieldCodeText(textContent).process(chr);
    } else {
      return textContent;
    }
  }
}

class OfficeMathFieldCodeText {
  constructor(mathText) {
    this.mathText = mathText;
  }

  process(chr) {
    const cancelMatch1 = this.mathText.match(/eq \\o\s?\((.*?),\/\)/i);
    const cancelMatch2 = this.mathText.match(/eq \\o\s?\((.*?),／\)/i);
    const overlineMatch = this.mathText.match(/eq \\x\s?\\to \((.*?)\)/i);

    if (cancelMatch1) {
      return `\\cancel{${cancelMatch1[1]}}`;
    }
    if (cancelMatch2) {
      return `\\cancel{${cancelMatch2[1]}}`;
    }
    if (overlineMatch) {
      return `\\overline{${overlineMatch[1]}}`;
    }

    return this.mathText;
  }
}
