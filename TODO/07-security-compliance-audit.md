# 07. Security & Compliance Audit (LOW)

## What?

Implement security and compliance features as specified in Section 8 of the integration spec: tool call logging for audit, OpenAI data logging controls, and enhanced error handling.

## Where?

- **New files**: `services/auditService.ts`, `services/securityService.ts`
- **Updates**: `app/api/chat/route.ts`, `middleware.ts` (new)
- **Configuration**: Environment variables and logging

## How?

### 7.1 Create AuditService

```typescript
// services/auditService.ts
interface AuditLogEntry {
  timestamp: string;
  userId: string;
  sessionId: string;
  action: string;
  toolName?: string;
  parameters?: any;
  result?: any;
  error?: string;
  duration: number;
  ipAddress?: string;
  userAgent?: string;
}

interface AuditConfig {
  enabled: boolean;
  logLevel: "basic" | "detailed" | "full";
  retentionDays: number;
  sensitiveFields: string[];
}

export class AuditService {
  private static instance: AuditService;
  private config: AuditConfig;
  private logs: AuditLogEntry[] = [];

  constructor() {
    this.config = {
      enabled: process.env.AUDIT_LOGGING_ENABLED === "true",
      logLevel: (process.env.AUDIT_LOG_LEVEL as any) || "basic",
      retentionDays: parseInt(process.env.AUDIT_RETENTION_DAYS || "30"),
      sensitiveFields: ["api_key", "password", "token"],
    };
  }

  static getInstance(): AuditService {
    if (!AuditService.instance) {
      AuditService.instance = new AuditService();
    }
    return AuditService.instance;
  }

  async logToolCall(
    entry: Omit<AuditLogEntry, "timestamp" | "duration">
  ): Promise<void> {
    if (!this.config.enabled) return;

    const auditEntry: AuditLogEntry = {
      ...entry,
      timestamp: new Date().toISOString(),
      duration: 0, // Will be calculated
    };

    // Sanitize sensitive data
    if (this.config.logLevel !== "full") {
      auditEntry.parameters = this.sanitizeSensitiveData(entry.parameters);
      auditEntry.result = this.sanitizeSensitiveData(entry.result);
    }

    this.logs.push(auditEntry);

    // Clean old logs
    await this.cleanOldLogs();
  }

  private sanitizeSensitiveData(data: any): any {
    if (!data || typeof data !== "object") return data;

    const sanitized = { ...data };
    this.config.sensitiveFields.forEach((field) => {
      if (sanitized[field]) {
        sanitized[field] = "[REDACTED]";
      }
    });

    return sanitized;
  }

  private async cleanOldLogs(): Promise<void> {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - this.config.retentionDays);

    this.logs = this.logs.filter((log) => new Date(log.timestamp) > cutoffDate);
  }

  async getAuditLogs(filters?: {
    userId?: string;
    toolName?: string;
    startDate?: string;
    endDate?: string;
  }): Promise<AuditLogEntry[]> {
    let filteredLogs = [...this.logs];

    if (filters?.userId) {
      filteredLogs = filteredLogs.filter(
        (log) => log.userId === filters.userId
      );
    }

    if (filters?.toolName) {
      filteredLogs = filteredLogs.filter(
        (log) => log.toolName === filters.toolName
      );
    }

    if (filters?.startDate) {
      filteredLogs = filteredLogs.filter(
        (log) => log.timestamp >= filters.startDate!
      );
    }

    if (filters?.endDate) {
      filteredLogs = filteredLogs.filter(
        (log) => log.timestamp <= filters.endDate!
      );
    }

    return filteredLogs;
  }
}
```

### 7.2 Create SecurityService

```typescript
// services/securityService.ts
interface SecurityConfig {
  enableDataLogging: boolean;
  maxRequestSize: number;
  rateLimitPerMinute: number;
  allowedOrigins: string[];
}

export class SecurityService {
  private static instance: SecurityService;
  private config: SecurityConfig;
  private requestCounts: Map<string, number[]> = new Map();

  constructor() {
    this.config = {
      enableDataLogging: process.env.OPENAI_DATA_LOGGING_ENABLED === "true",
      maxRequestSize: parseInt(process.env.MAX_REQUEST_SIZE || "1048576"), // 1MB
      rateLimitPerMinute: parseInt(process.env.RATE_LIMIT_PER_MINUTE || "60"),
      allowedOrigins: process.env.ALLOWED_ORIGINS?.split(",") || ["*"],
    };
  }

  static getInstance(): SecurityService {
    if (!SecurityService.instance) {
      SecurityService.instance = new SecurityService();
    }
    return SecurityService.instance;
  }

  async validateRequest(
    req: Request
  ): Promise<{ valid: boolean; error?: string }> {
    // Check request size
    const contentLength = req.headers.get("content-length");
    if (contentLength && parseInt(contentLength) > this.config.maxRequestSize) {
      return { valid: false, error: "Request too large" };
    }

    // Check rate limiting
    const clientId = this.getClientId(req);
    if (!this.checkRateLimit(clientId)) {
      return { valid: false, error: "Rate limit exceeded" };
    }

    // Check origin
    const origin = req.headers.get("origin");
    if (origin && !this.isAllowedOrigin(origin)) {
      return { valid: false, error: "Origin not allowed" };
    }

    return { valid: true };
  }

  private getClientId(req: Request): string {
    // Use IP address or user ID as client identifier
    return (
      req.headers.get("x-forwarded-for") ||
      req.headers.get("x-real-ip") ||
      "unknown"
    );
  }

  private checkRateLimit(clientId: string): boolean {
    const now = Date.now();
    const minuteAgo = now - 60000;

    if (!this.requestCounts.has(clientId)) {
      this.requestCounts.set(clientId, []);
    }

    const requests = this.requestCounts.get(clientId)!;

    // Remove old requests
    const recentRequests = requests.filter((time) => time > minuteAgo);
    this.requestCounts.set(clientId, recentRequests);

    // Check if under limit
    if (recentRequests.length >= this.config.rateLimitPerMinute) {
      return false;
    }

    // Add current request
    recentRequests.push(now);
    return true;
  }

  private isAllowedOrigin(origin: string): boolean {
    return (
      this.config.allowedOrigins.includes("*") ||
      this.config.allowedOrigins.includes(origin)
    );
  }

  shouldLogData(): boolean {
    return this.config.enableDataLogging;
  }
}
```

### 7.3 Update Chat Route with Security

```typescript
// In app/api/chat/route.ts
import { AuditService } from "@/services/auditService";
import { SecurityService } from "@/services/securityService";

export async function POST(req: Request) {
  const startTime = Date.now();
  const auditService = AuditService.getInstance();
  const securityService = SecurityService.getInstance();

  try {
    // Validate request
    const validation = await securityService.validateRequest(req);
    if (!validation.valid) {
      return new Response(JSON.stringify({ error: validation.error }), {
        status: 429,
        headers: { "Content-Type": "application/json" },
      });
    }

    const { messages } = await req.json();

    // Log tool calls with audit
    const result = streamText({
      model: openai("gpt-4o-mini"),
      system: await generateSchemaAwareSystemPrompt(),
      messages,
      temperature: 0.1,
      maxTokens: 4096,
      maxSteps: 10,
      tools: {
        // All tools with audit logging
        list_columns: tool({
          description: "Get column names and row count",
          parameters: z.object({}),
          execute: async () => {
            const start = Date.now();
            try {
              const result =
                await UniverService.getInstance().getSheetContext();
              await auditService.logToolCall({
                userId: "user", // Get from auth
                sessionId: "session", // Get from session
                action: "tool_execution",
                toolName: "list_columns",
                parameters: {},
                result: { columns: result.tables[0]?.headers || [] },
                duration: Date.now() - start,
                ipAddress: req.headers.get("x-forwarded-for") || "unknown",
              });
              return result;
            } catch (error) {
              await auditService.logToolCall({
                userId: "user",
                sessionId: "session",
                action: "tool_error",
                toolName: "list_columns",
                parameters: {},
                error: error instanceof Error ? error.message : "Unknown error",
                duration: Date.now() - start,
                ipAddress: req.headers.get("x-forwarded-for") || "unknown",
              });
              throw error;
            }
          },
        }),
        // ... other tools with similar audit logging
      },
    });

    return result.toDataStreamResponse();
  } catch (error) {
    const duration = Date.now() - startTime;
    await auditService.logToolCall({
      userId: "user",
      sessionId: "session",
      action: "request_error",
      error: error instanceof Error ? error.message : "Unknown error",
      duration,
      ipAddress: req.headers.get("x-forwarded-for") || "unknown",
    });

    return new Response(JSON.stringify({ error: "Internal server error" }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
}
```

### 7.4 Add Environment Variables

```bash
# .env.local
AUDIT_LOGGING_ENABLED=true
AUDIT_LOG_LEVEL=detailed
AUDIT_RETENTION_DAYS=30
OPENAI_DATA_LOGGING_ENABLED=false
MAX_REQUEST_SIZE=1048576
RATE_LIMIT_PER_MINUTE=60
ALLOWED_ORIGINS=http://localhost:3000,https://yourdomain.com
```

## How to Test?

1. **Audit Logging Tests**:

   ```typescript
   describe("AuditService", () => {
     it("should log tool calls", async () => {
       const service = AuditService.getInstance();
       await service.logToolCall({
         userId: "test-user",
         sessionId: "test-session",
         action: "tool_execution",
         toolName: "list_columns",
       });

       const logs = await service.getAuditLogs({ userId: "test-user" });
       expect(logs).toHaveLength(1);
       expect(logs[0].toolName).toBe("list_columns");
     });
   });
   ```

2. **Security Tests**:

   - Test rate limiting
   - Test request size validation
   - Test origin validation

3. **Integration Tests**:
   - Test audit logging in real requests
   - Verify sensitive data redaction
   - Test error logging

## Important Dependencies to Not Break

- **Existing tool execution** - Don't break current functionality
- **Performance** - Don't significantly slow down requests
- **User experience** - Don't block legitimate requests
- **Data privacy** - Ensure proper data handling

## Dependencies That Will Work Thanks to This

- **Compliance requirements** - Meet audit and security standards
- **Production readiness** - Enterprise-grade security
- **Monitoring** - Track usage and errors
- **Debugging** - Better error tracking

## Implementation Strategy (Least Disruptive)

1. **Start with audit logging** - Add logging without breaking functionality
2. **Implement security checks** - Add validation gradually
3. **Test thoroughly** - Ensure no performance impact
4. **Add configuration** - Make features configurable
5. **Document security** - Provide security documentation

## Priority Order

1. Create AuditService with basic logging
2. Add security validation to chat route
3. Implement rate limiting
4. Add sensitive data redaction
5. Create security configuration
6. Test with real requests
7. Add monitoring and alerting
