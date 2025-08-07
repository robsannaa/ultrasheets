# UltraSheets

## 🎥 Demo

[Watch UltraSheets Demo](https://drive.google.com/file/d/16uO7e7AuWmeSh9OQ8DlWah7CaIVQvv1N/view?usp=sharing)

> **Note**: Click the link above to watch the demo video on Google Drive.

An AI-powered spreadsheet application built with Next.js, Univer, and OpenAI. This application provides intelligent spreadsheet analysis, data manipulation, and visualization capabilities through natural language interactions.

## 🚀 Features

- **AI-Powered Chat Interface**: Natural language spreadsheet operations
- **Native Univer Integration**: Full-featured spreadsheet engine
- **Real-time Data Analysis**: Automatic detection of data structures, headers, and numeric columns
- **Smart Chart Generation**: Create native Univer charts with AI assistance
- **Pivot Table Creation**: Automated pivot table generation and analysis
- **Multi-Sheet Support**: Work across multiple sheets in a workbook
- **Financial Intelligence**: Advanced financial calculations and formatting

## 🛠 Tech Stack

- **Frontend**: Next.js 15, React, TypeScript
- **Spreadsheet Engine**: Univer (UniverJS)
- **AI Integration**: OpenAI GPT-4, Vercel AI SDK
- **Styling**: Tailwind CSS, shadcn/ui
- **State Management**: React hooks, Zustand-like patterns

## 📋 Prerequisites

- Node.js 18+
- npm or pnpm
- OpenAI API key

## 🔧 Setup Instructions

1. **Clone the repository**

   ```bash
   git clone <repository-url>
   cd ultrasheets
   ```

2. **Install dependencies**

   ```bash
   npm install
   # or
   pnpm install
   ```

3. **Environment Configuration**
   Create a `.env.local` file in the root directory:

   ```env
   OPENAI_API_KEY=your_openai_api_key_here
   ```

4. **Run the development server**

   ```bash
   npm run dev
   # or
   pnpm dev
   ```

5. **Open your browser**
   Navigate to `http://localhost:3000`

## 🏗 Architecture Overview

### Core Components

- **`/app/page.tsx`** - Main application layout with resizable panels
- **`/components/univer.tsx`** - Univer spreadsheet engine integration
- **`/components/chat-sidebar.tsx`** - AI chat interface and tool execution
- **`/app/api/chat/route.ts`** - LLM API endpoint with tool definitions

### Services

- **`/services/univerService.ts`** - Univer instance management and spreadsheet operations
- **`/services/llmSpreadsheetController.ts`** - Natural language processing for spreadsheet commands
- **`/services/chartService.ts`** - Chart generation and management
- **`/services/pivotTableService.ts`** - Pivot table operations
- **`/services/financialAnalysisService.ts`** - Financial calculations and analysis

### Key Patterns

1. **Component Communication**: Global `window.univerAPI` for cross-component access
2. **Service Layer**: Singleton pattern for Univer service management
3. **Tool System**: Server-side tool execution with client-side application
4. **Data Flow**: Chat → API → Tools → Client Actions → Univer Updates

## 🔍 Known Issues & Solutions

### Issue 1: UniverService Not Ready

**Problem**: `UniverService.isReady()` returns false, preventing LLM from accessing spreadsheet data.

**Root Cause**: Timing issue between component initialization and service readiness checks.

**Solution**: Modified `extractWorkbookData()` in `chat-sidebar.tsx` to:

- Bypass service readiness check
- Use direct `window.univerAPI` access with UniverService fallback
- Implemented dual-path data extraction

### Issue 2: Missing Helper Functions in API Route

**Problem**: `ReferenceError: getActiveSheet is not defined` in chart generation and other tools.

**Root Cause**: Missing helper functions referenced by tool implementations.

**Solution**: Added complete set of helper functions in `/app/api/chat/route.ts`:

- `getActiveSheet()`
- `extractHeadersFromSheet()`
- `countDataRowsInSheet()`
- `getAllSheetNames()`
- `getSheetByName()`
- And others for comprehensive tool support

## 🛡 Troubleshooting

### Development Issues

1. **Univer not loading**: Check browser console for initialization errors
2. **Chat not responding**: Verify OpenAI API key in `.env.local`
3. **Tools failing**: Check server logs for missing function errors
4. **Data not extracting**: Ensure `window.univerAPI` is set in browser dev tools

### Common Error Messages

- `"Univer instance not ready"` → Wait for component to fully load or refresh page
- `"No workbook data available"` → Ensure spreadsheet has data in cells
- `"Column not found"` → Check column names match data headers
- `"getActiveSheet is not defined"` → Fixed in current version

### Debug Mode

Enable detailed logging by checking browser console for:

- `🚀 Univer component: Starting initialization...`
- `📊 extractWorkbookData: Found X sheets in workbook`
- `✅ Successfully set cell...`

## 📚 Usage Examples

### Basic Operations

```
User: "List the columns in this spreadsheet"
AI: Lists all column headers with data counts

User: "Calculate the total of the Sales column"
AI: Adds SUM formula to spreadsheet automatically

User: "Create a chart showing sales by region"
AI: Generates native Univer chart with proper data range
```

### Advanced Features

```
User: "Create a pivot table grouping by Date with sum of Revenue"
AI: Creates formatted pivot table in spreadsheet

User: "Format the headers as bold and color them blue"
AI: Applies formatting directly to spreadsheet cells

User: "Switch to the P&L sheet and analyze the data"
AI: Switches sheets and provides comprehensive analysis
```

## 🔮 Future Enhancements

- [ ] Multiple file upload support
- [ ] Enhanced chart customization
- [ ] Advanced financial modeling templates
- [ ] Real-time collaboration features
- [ ] Export to Excel/CSV functionality
- [ ] Custom formula suggestions
- [ ] Data validation and error checking

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues and questions:

1. Check the troubleshooting section above
2. Review browser console for error messages
3. Ensure all dependencies are properly installed
4. Verify OpenAI API key is correctly configured
