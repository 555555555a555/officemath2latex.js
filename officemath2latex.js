//
// officemath2latex.js
//
// This program defines a function named processMathNode that converts mathematical
// expressions in OfficeMath format (DocumentFormat.OpenXml.Math) to LaTeX code.
// It takes an XML element as input and returns the corresponding LaTeX representation.
// The input mathXml element must be "m:oMath" node.
//

function processMathNode(mathXml, chr = "") {
  const mathNode = new OfficeMathNode(mathXml);
  return mathNode.process();

  //
  return mathNode
    .process()
    .replace(/\\ /g, "\\spAce")
    .replace(/ \\/g, "\\")
    .replace(/\\spAce/g, "\\ ");
}

class OfficeMathNode {
  constructor(node) {
    this.node = node;
  }
  process(chr = "") {
    let mathString = "";
    for (const element of this.node.childNodes) {
      const childElement = new OfficeMathElement(element);
      mathString += childElement.process(chr);
    }
    return mathString;
  }
}

class OfficeMathElement {
  constructor(node) {
    this.node = node;
  }

  process(chr = "") {
    switch (this.node.nodeName) {
      case "m:e":
        return new OfficeMathNode(this.node).process(chr) + chr;
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
      case "m:limUpp":
        return new OfficeMathLimUpper(this.node).process(chr);
      case "m:m":
        return new OfficeMathMatrix(this.node).process(chr);
      case "m:acc":
        return new OfficeMathAccent(this.node).process(chr);
      case "m:groupChr":
        return new OfficeMathGroupCharacter(this.node).process(chr);
      case "m:bar":
        return new OfficeMathBar(this.node).process(chr);
      case "m:rad":
        return new OfficeMathRadical(this.node).process(chr);
      case "m:eqArr":
        return new OfficeMathEquationArray(this.node).process(chr);
      case "m:box":
        return `{${new OfficeMathNode(this.node).process(chr)}}`;
      case "m:borderBox":
        return `\\boxed{${new OfficeMathNode(this.node).process(chr)}}`;
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
        mathString += new OfficeMathNode(mathElement).process(chr);
      } else if (mathElement.nodeName === "m:sup" && type === "^") {
        mathString += `^{${new OfficeMathNode(mathElement).process(chr)}}`;
      } else if (mathElement.nodeName === "m:sub" && type === "_") {
        mathString += `_{${new OfficeMathNode(mathElement).process(chr)}}`;
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
        mathString += new OfficeMathNode(mathElement).process(chr);
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
        openPar = this.getOpenBracket(val);
      } else if (dPr.nodeName === "m:endChr") {
        const val = dPr.getAttribute("m:val");
        closePar = this.getCloseBracket(val);
      }
    }

    return { openPar, closePar };
  }

  getOpenBracket(val) {
    switch (val) {
      case "|":
        return "\\left| ";
      case "{":
        return "\\left\\{ ";
      case "[":
        return "\\left\\lbrack ";
      case "]":
        return "\\left\\rbrack ";
      case "〈":
        return "\\left\\langle ";
      case "⌊":
        return "\\left\\lfloor ";
      case "⌈":
        return "\\left\\lceil ";
      case "‖":
        return "\\left\\| ";
      case "⟦":
        return "\\left. ⟦ ";
      default:
        return "\\left. ";
    }
  }

  getCloseBracket(val) {
    switch (val) {
      case "|":
        return "\\right| ";
      case "}":
        return "\\right\\} ";
      case "[":
        return "\\right\\lbrack ";
      case "]":
        return "\\right\\rbrack ";
      case "〉":
        return "\\right\\rangle ";
      case "⌋":
        return "\\right\\rfloor ";
      case "⌉":
        return "\\right\\rceil ";
      case "‖":
        return "\\right\\| ";
      case "⟧":
        return "\\right.⟧ ";
      default:
        return "\\right. ";
    }
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
        numString = new OfficeMathNode(mathElement).process(chr);
      } else if (mathElement.nodeName === "m:den") {
        denString = new OfficeMathNode(mathElement).process(chr);
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
        subString = `_{${new OfficeMathNode(mathElement).process(chr)}}`;
      } else if (mathElement.nodeName === "m:sup") {
        supString = `^{${new OfficeMathNode(mathElement).process(chr)}}`;
      } else {
        postString += `{${new OfficeMathNode(mathElement).process(chr)}}`;
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
        const functionName = new OfficeMathNode(mathElement).process(chr);
        mathString += this.getFunctionString(functionName);
      } else {
        mathString += new OfficeMathNode(mathElement).process(chr);
      }
    }
    return `${mathString}}`;
  }

  getFunctionString(functionName) {
    switch (functionName) {
      case "tan":
      case "csc":
      case "sec":
      case "cot":
      case "sin":
      case "cos":
      case "tan":
      case "csc":
      case "sec":
      case "cot":
      case "sinh":
      case "cosh":
      case "tanh":
      case "coth":
        //case "csch": // not defined in LaTeX
        //case "sech": // not defined in LaTeX
        return `\\${functionName}{`;
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
        const elementString = new OfficeMathNode(mathElement).process(chr);
        mathString += elementString.trim() === "lim" ? `\\${elementString.trim()}` : elementString;
      } else if (mathElement.nodeName === "m:lim") {
        mathString += `_{${new OfficeMathNode(mathElement).process(chr)}}`;
      }
    }
    return mathString;
  }
}

class OfficeMathLimUpper extends OfficeMathElement {
  process(chr) {
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:e") {
        const elementString = new OfficeMathNode(mathElement).process(chr);
        mathString += elementString.trim() === "lim" ? `\\${elementString.trim()}` : elementString;
      } else if (mathElement.nodeName === "m:lim") {
        mathString += `^{${new OfficeMathNode(mathElement).process(chr)}}`;
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
        mathString += `${new OfficeMathNode(mathElement).process("&").slice(0, -1)} \\\\ \n`;
      }
    }
    return `${mathString}\\end{matrix} `;
  }
}

class OfficeMathAccent extends OfficeMathElement {
  process(chr) {
    let accentString = "\\widehat{";
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:accPr") {
        for (const naryPr of mathElement.childNodes) {
          if (naryPr.nodeName === "m:chr") {
            const val = naryPr.getAttribute("m:val");
            accentString = this.getAccentOperator(val);
          }
        }
      }
      if (mathElement.nodeName === "m:e") {
        mathString += new OfficeMathNode(mathElement).process(chr);
      }
    }
    return `${accentString}${mathString}${"}".repeat(
      accentString.split("{").length - 1 - (accentString.split("}").length - 1)
    )}`;
  }
  getAccentOperator(val) {
    switch (val) {
      case "̇":
        return "\\dot{";
      case "̈":
        return "\\ddot{";
      case "⃛":
        return "\\dddot{";
      case "̌":
        return "\\check{";
      case "́":
        return "\\acute{";
      case "̀":
        return "\\grave{";
      case "̆":
        return "\\breve{";
      case "̃":
        return "\\widetilde{";
      case "̅":
        return "\\overline{";
      case "̿":
        return "\\overline{\\overline{";
      case "⃖":
        return "\\overleftarrow{";
      case "⃡":
        return "\\overleftrightarrow{";
      case "⃐":
        return "\\overset{\\leftharpoonup}{";
      case "⃑":
        return "\\overset{\\rightharpoonup}{";

      //"̀"
      default:
        return "\\overrightarrow{";
    }
  }
}

class OfficeMathGroupCharacter extends OfficeMathElement {
  //accentPrefix = "\\overset{";
  //accentPostfix = "}{︸}";
  accentPrefix = "\\underbrace{";
  accentPostfix = "}";

  process(chr) {
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:groupChrPr") {
        for (const naryPr of mathElement.childNodes) {
          if (naryPr.nodeName === "m:chr") {
            const val = naryPr.getAttribute("m:val");
            this.accentPrefix = this.getAccentOperator(val);
          }
        }
      }
      if (mathElement.nodeName === "m:e") {
        mathString += new OfficeMathNode(mathElement).process(chr);
      }
    }
    return `${this.accentPrefix}${mathString}${this.accentPostfix}`;
  }
  getAccentOperator(val) {
    switch (val) {
      case "⏞": // <m:pos m:val="top"/><m:vertJc m:val="bot"/>
        this.accentPostfix = "}";
        return "\\overbrace{";
      //"̀"
      default:
        this.accentPostfix = "}";
        return "\\GCHR2{";
    }
  }
}

class OfficeMathBar extends OfficeMathElement {
  process(chr) {
    let accentString = "\\underline{";
    let mathString = "";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:barPr") {
        for (const barPr of mathElement.childNodes) {
          if (barPr.nodeName === "m:pos") {
            const val = barPr.getAttribute("m:val");
            if (val === "top") accentString = "\\overline{";
          }
        }
      }
      if (mathElement.nodeName === "m:e") {
        mathString += new OfficeMathNode(mathElement).process(chr);
      }
    }
    return `${accentString}${mathString}${"}".repeat(
      accentString.split("{").length - 1 - (accentString.split("}").length - 1)
    )}`;
  }
}

class OfficeMathRadical extends OfficeMathElement {
  process(chr) {
    let mathString = "\\sqrt";
    for (const mathElement of this.node.childNodes) {
      if (mathElement.nodeName === "m:deg") {
        mathString += `[${new OfficeMathNode(mathElement).process(chr)}]`;
      } else if (mathElement.nodeName === "m:e") {
        mathString += `{${new OfficeMathNode(mathElement).process(chr)}}`;
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
        mathString += `${new OfficeMathNode(mathElement).process(chr)} \\\\ \n`;
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
        mathString += `_{${new OfficeMathNode(mathElement).process(chr)}}`;
      } else if (mathElement.nodeName === "m:sup") {
        mathString += `^{${new OfficeMathNode(mathElement).process(chr)}}`;
      } else if (mathElement.nodeName !== "m:ctrlPr") {
        mathString += new OfficeMathNode(mathElement).process(chr);
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
      return `\\cancel{${cancelMatch1[1]}}`; // add \usepackage{cancel}
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
