# officemath2latex.js

This program defines a function named `processMathNode` that converts mathematical
expressions in OfficeMath format (DocumentFormat.OpenXml.Math) to LaTeX code.
It takes an XML element as input and returns the corresponding LaTeX representation.
The input mathXml element must be `m:oMath` node.

It handles all formulas included as samples in MS Word equations and any mathematical 
elements you can create using the Word equation toolbar. This should be sufficient for 
most users. However, we understand that some users might require advanced functionalities 
for complex equations.

Don't hesitate to let us know if you encounter any issues or have suggestions for improvement!

## How to use

For detailed usage instructions and examples, please refer to the `officemath2latex_sample.html` file provided in the repository.
This file demonstrates how to integrate the function into your project.

You can access directly to the example.
https://555555555a555.github.io/officemath2latex.js/officemath2latex_sample.html

## License
This project is licensed under the MIT License.
