import { ZodRawShape, z } from 'zod';

type ToolResult = { content: Array<{ type: string; text: string }>; isError?: boolean };
type Handler = (args: Record<string, unknown>) => Promise<ToolResult>;

/**
 * Lightweight stand-in for McpServer that captures tool registrations
 * so individual handlers can be unit-tested without starting a real server.
 */
export class MockMcpServer {
  private registry = new Map<string, { schema: ZodRawShape; handler: Handler }>();

  tool<T extends ZodRawShape>(
    name: string,
    _description: string,
    schema: T,
    handler: (args: z.infer<z.ZodObject<T>>) => Promise<ToolResult>
  ): void {
    this.registry.set(name, { schema, handler: handler as Handler });
  }

  /** Validate args through the Zod schema, then invoke the handler. */
  async call(name: string, rawArgs: Record<string, unknown>): Promise<ToolResult> {
    const entry = this.registry.get(name);
    if (!entry) throw new Error(`Tool "${name}" is not registered`);
    const parsed = z.object(entry.schema).parse(rawArgs);
    return entry.handler(parsed as Record<string, unknown>);
  }

  isRegistered(name: string): boolean {
    return this.registry.has(name);
  }

  registeredNames(): string[] {
    return [...this.registry.keys()];
  }
}
