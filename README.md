# DocxWeaver

## Overview
DocxWeaver is a Python class designed to convert, translate, and modify Word documents in various ways, such as adding comments, transforming text, and both. It's suitable for automating document workflows where changes to the document based on dynamic inputs are required.

## Features
- **Comments Only**: Adds comments to the document without altering the text.
- **Transform Only**: Modifies the document's text in-place without adding comments.
- **Transform and Comments**: Transforms the document's text and retains the original text in comments for reference.

## Requirements
- Available in Dockerfile

## Installation
To use DocxWeaver, you can build the container use make build-docker-image, or work in the .devcontainer provided.

## Usage
```python
from weaver.weaver import DocxWeaver

# Comment Only
doc = DocxWeaver(
    filename="fake-consulting-doc.docx",
    purpose="You are reviewing a consulting agreement, from the perspective of the consultant.",
    paragraph_prompt="Review this and highlight any issues or concerns.",
    table_prompt=None,
    mode="comments_only",
)
weave_result = doc.weave_document(output_fn="fake-consulting-doc-review.docx")

# Transform Only
doc = DocxWeaver(
    filename="fake-consulting-doc.docx",
    purpose="You are converting a consulting document into one that rhymes.",
    paragraph_prompt="Convert the following paragraph into a rhyming version.",
    table_prompt=None,
    mode="transform_only",
)
weave_result = doc.weave_document(output_fn="fake-consulting-doc-transform.docx")

# Transform And Comments
doc = DocxWeaver(
    filename="fake-consulting-doc.docx",
    purpose="You are translating a consulting document into french.",
    paragraph_prompt="Convert the following paragraph into french.",
    table_prompt="Convert the following table cell into french",
    mode="transform_and_comments",
)
weave_result = doc.weave_document(output_fn="fake-consulting-doc-transform-comments.docx")
```

## Documentation
For further details, refer to the inline comments in the DocxWeaver class definition. Each method and its parameters are documented to explain their purpose and usage.

## Contributing
Contributions to improve DocxWeaver are welcome. Please feel free to fork the repository, make your changes, and submit a pull request.

## Licensing

DocxWeaver is dual-licensed:

1. **Open Source License**: For open source projects and individual use, DocxWeaver is available under the MIT License. Under this license, you are free to use, modify, and distribute the software as part of your open source projects with the requirement to include the original copyright and license notice in any copy of the software/source code.

2. **Commercial License**: For commercial use, including incorporation of DocxWeaver into proprietary software or as part of a commercial product or service, a commercial license is required. The commercial license grants you, the licensee, the rights to develop, market, and distribute your product or service using DocxWeaver subject to the payment of a licensing fee and adherence to the terms specified in the commercial license agreement.

Please contact Jesse Moore for more details about obtaining a commercial license.

By using, copying, modifying, or distributing the software (or any work based on the software), you agree to the terms of these licenses. If you do not agree to the terms of these licenses, do not use, copy, modify, or distribute the software.
