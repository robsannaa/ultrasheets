import { registerOTel } from "@vercel/otel";
import { initializeOTEL } from "langsmith/experimental/otel/setup";

const { DEFAULT_LANGSMITH_SPAN_PROCESSOR } = initializeOTEL({});

export function register() {
  registerOTel({
    serviceName: "your-project-name",
    spanProcessors: [DEFAULT_LANGSMITH_SPAN_PROCESSOR],
  });
}
