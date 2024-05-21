# officemath2latex.js

This program defines a function named `processMathNode` that converts mathematical
expressions in OfficeMath format (DocumentFormat.OpenXml.Math) to LaTeX code.
It takes an XML element as input and returns the corresponding LaTeX representation.
The input mathXml element must be `m:oMath` node.

It can convert most of the formulas that are included as samples in MS Word equations, but I think this will be enough in most cases.

## HOW TO USE:

For detailed usage instructions and examples, please refer to the `officemath2latex_sample.html` file provided in the repository.
This file demonstrates how to integrate the function into your project.

## MIT License
This project is licensed under the MIT License.
