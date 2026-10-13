"""xlsx-agent: workbooks as a source an agent can describe, query and derive from.

The agent never writes XML and never does arithmetic: it reads a catalogue of the workbook,
asks for rows, and asks the library for figures, each of which says which cells it came
from.  Built on lxml, ooxml-edit and ooxml-common, beside pptx-agent and docx-agent.

Nothing is built yet beyond the package itself; see ROADMAP.md for the plan and what each
milestone adds.
"""

__version__ = "0.0.1"
