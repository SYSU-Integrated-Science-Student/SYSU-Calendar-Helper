# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a calendar-helper tool that converts Word timetable exports (Flat OPC XML format) into Apple Calendar ICS files. The project includes both a CLI tool and a web frontend for parsing course schedules from Chinese university systems.

## Package Management

- **Package Manager**: pnpm (specified in package.json)
- **Node.js**: Requires Node.js >= 18
- **Type**: ES Modules (type: "module" in package.json)

## Development Commands

### Core Scripts
- `pnpm generate:ics` - Generate ICS file from sample data
- `pnpm generate:json` - Generate JSON format output from sample data
- `pnpm frontend:serve` - Start the web frontend server
- `pnpm frontend:serve:open` - Start server and open browser
- `pnpm check` - Syntax check core JavaScript files

### Binary Commands
- `generate-ics` - CLI tool for ICS generation (accepts --input, --start, --tz, --format flags)
- `serve-frontend` - Web server for the frontend (accepts --port, --host, --open flags)

## Architecture

### Core Components

**src/schedule.js** - Main parsing engine:
- Parses Word XML documents (Flat OPC or ZIP format)
- Extracts course information from table structures
- Handles Chinese day names (星期一 to 星期日)
- Processes course scheduling with week patterns
- Generates ICS calendar events with proper timezone handling

**bin/generate-ics.js** - CLI interface:
- Command-line argument parsing
- File input handling
- Output format selection (ICS/JSON)
- Error handling and help display

**bin/serve-frontend.js** - Web server:
- Static file serving for frontend
- API endpoint `/api/parse` for file processing
- CORS support for cross-origin requests
- Browser auto-opening option

**frontend/** - Web interface:
- File upload with drag-and-drop support
- Date and timezone configuration
- Real-time parsing feedback
- ICS file download functionality

### Data Flow
1. Input: Word document (Flat OPC XML or .docx)
2. Parse: Extract document.xml and parse table structure
3. Process: Map Chinese day names, extract course details, handle week patterns
4. Generate: Create calendar events with proper recurrence rules
5. Output: ICS file format or JSON data

### Key Features
- Supports both Flat OPC XML and .docx file formats
- Handles Chinese university course schedule formats
- Proper timezone support for international students
- RRULE generation for recurring events
- Web interface with real-time parsing

## File Formats

### Input Formats
- Flat OPC XML (sample-package.xml)
- Word .docx files (ZIP format)

### Output Formats
- ICS (iCalendar format) for Apple Calendar import
- JSON for debugging and data inspection

## Sample Data

Located in `samples/` directory:
- `sample-package.xml` - Flat OPC XML example
- `sample-docx.docx` - Word document example
- `generated-document.xml` - Generated output example

## Error Handling

The system provides detailed error messages for:
- Unsupported file formats
- Missing required data in schedules
- Timezone parsing issues
- XML parsing failures
- ZIP extraction problems