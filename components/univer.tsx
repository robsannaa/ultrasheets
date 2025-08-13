"use client";

import { useEffect, useRef } from "react";
import { useTheme } from "next-themes";
// CSS imports are safe to keep at module level
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-drawing/lib/index.css";
import "@univerjs/preset-sheets-advanced/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";
import "@univerjs/preset-sheets-table/lib/index.css";

// Global tool execution handler
declare global {
  interface Window {
    executeUniverTool: (toolName: string, params?: any) => Promise<any>;
    lastSumCell?: string; // Track the last sum cell created
    __ultraUniverInitialized?: boolean;
    __ultraUniverAPI?: any;
    __ultraUniverLifecycleHookAdded?: boolean;
    __ultraUniverDispose?: () => void;
    // no watermark observer
  }
}

export function Univer() {
  const containerRef = useRef<HTMLDivElement>(null!);
  const { resolvedTheme } = useTheme();

  useEffect(() => {
    // Suppress Univer's duplicate identifier warnings during development
    const originalError = console.error;
    const originalWarn = console.warn;

    // Do not override console.error to keep React error stacks accurate

    console.warn = (...args) => {
      const message = args.join(" ");
      if (
        message.includes("Identifier") &&
        message.includes("already exists")
      ) {
        return; // Suppress these specific warnings
      }
      originalWarn.apply(console, args);
    };

    // Dynamic imports to avoid SSR issues
    const initializeUniver = async () => {
      // IMPORTANT: Preset mode only. Do NOT import facades in preset mode,
      // as it double-registers plugins and breaks DI (redi errors).
      const { UniverSheetsCorePreset } = await import(
        "@univerjs/preset-sheets-core"
      );
      const { UniverSheetsDrawingPreset } = await import(
        "@univerjs/preset-sheets-drawing"
      );
      const { UniverSheetsAdvancedPreset } = await import(
        "@univerjs/preset-sheets-advanced"
      );
      const { UniverSheetsConditionalFormattingPreset } = await import(
        "@univerjs/preset-sheets-conditional-formatting"
      );

      // Table preset (optional; fall back if missing)
      let TablePreset: any = null;
      let tableLocales: any = null;
      try {
        const tablePresetMod = await import("@univerjs/preset-sheets-table");
        const tableLocaleMod = await import(
          "@univerjs/preset-sheets-table/locales/en-US"
        );
        TablePreset = tablePresetMod.UniverSheetsTablePreset;
        tableLocales = tableLocaleMod.default;
      } catch (e) {
        console.warn("⚠️ Table preset unavailable, continuing without it.", e);
      }

      // Use CalculationMode from sheets-formula to align with facade API
      const { CalculationMode } = await import("@univerjs/sheets-formula");
      const sheetsCoreEnUS = await import(
        "@univerjs/preset-sheets-core/locales/en-US"
      ).then((m) => m.default);
      const sheetsDrawingEnUS = await import(
        "@univerjs/preset-sheets-drawing/locales/en-US"
      ).then((m) => m.default);
      const sheetsAdvancedEnUS = await import(
        "@univerjs/preset-sheets-advanced/locales/en-US"
      ).then((m) => m.default);
      const sheetsCFEnUS = await import(
        "@univerjs/preset-sheets-conditional-formatting/locales/en-US"
      ).then((m) => m.default);
      const {
        createUniver,
        LocaleType,
        mergeLocales,
        defaultTheme,
        greenTheme,
      } = await import("@univerjs/presets");
      // LifecycleStages import removed; no lifecycle hook needed

      // Enable filter and sort presets with proper error handling
      let filterPreset = null;
      let filterLocales = null;
      try {
        const filterModule = await import("@univerjs/preset-sheets-filter");
        const filterLocaleModule = await import(
          "@univerjs/preset-sheets-filter/locales/en-US"
        );
        filterPreset = filterModule.UniverSheetsFilterPreset;
        filterLocales = filterLocaleModule.default;
      } catch (error) {
        console.warn(
          "⚠️ Filter preset not available, continuing without filters:",
          error
        );
      }

      // Re-enable sort preset with proper error handling
      let sortPreset = null;
      let sortLocales = null;
      try {
        const sortModule = await import("@univerjs/preset-sheets-sort");
        const sortLocaleModule = await import(
          "@univerjs/preset-sheets-sort/locales/en-US"
        );
        sortPreset = sortModule.UniverSheetsSortPreset;
        sortLocales = sortLocaleModule.default;
      } catch (error) {
        console.warn(
          "⚠️ Sort preset not available, continuing without sorting:",
          error
        );
      }

      const corePreset = UniverSheetsCorePreset({
        container: containerRef.current,
        // Ensure formulas compute deterministically on load
        formula: {
          initialFormulaComputing: CalculationMode.FORCED,
        },
      });

      // Build presets in a way that allows graceful fallback if advanced/drawing cause injector errors
      let presets: any[] = [corePreset];
      try {
        // Try to include chart-related presets
        presets = [
          corePreset,
          UniverSheetsDrawingPreset(),
          UniverSheetsAdvancedPreset(),
          UniverSheetsConditionalFormattingPreset(),
        ];
      } catch (e) {
        console.warn(
          "⚠️ Failed to prepare chart presets; continuing with core only.",
          e
        );
        presets = [corePreset];
      }

      // Add filter preset if available
      if (filterPreset) {
        presets.push(filterPreset());
      }
      // Add sort preset if available
      if (sortPreset) {
        presets.push(sortPreset());
      }
      if (TablePreset) {
        presets.push(TablePreset());
      }

      let locales = mergeLocales(sheetsCoreEnUS);
      try {
        const additionalLocales = [
          ...(filterLocales ? [filterLocales] : []),
          ...(sortLocales ? [sortLocales] : []),
          ...(tableLocales ? [tableLocales] : []),
        ];

        locales = mergeLocales(
          sheetsCoreEnUS,
          sheetsDrawingEnUS,
          sheetsAdvancedEnUS,
          sheetsCFEnUS,
          ...additionalLocales
        );
      } catch (e) {
        console.warn(
          "⚠️ Failed to merge locales for chart presets; using core locales only.",
          e
        );
        const additionalLocales = [
          ...(filterLocales ? [filterLocales] : []),
          ...(sortLocales ? [sortLocales] : []),
          ...(tableLocales ? [tableLocales] : []),
        ];

        locales = mergeLocales(sheetsCoreEnUS, ...additionalLocales);
      }

      // Minimal watermark removal: call official API if present; no heuristics
      const removeUniverWatermark = (api: any) => {
        try {
          if (typeof api?.removeWatermark === "function") api.removeWatermark();
          else if (typeof api?.deleteWatermark === "function")
            api.deleteWatermark();
          else if (typeof api?.hideWatermark === "function")
            api.hideWatermark();
        } catch {}
      };

      // If already initialized (HMR / re-render), reuse existing API
      if (
        typeof window !== "undefined" &&
        window.__ultraUniverInitialized &&
        window.__ultraUniverAPI
      ) {
        // Remove Univer watermark if API available (HMR-safe)
        try {
          removeUniverWatermark(window.__ultraUniverAPI);
        } catch {}

        // Reuse existing API
        const univerAPI = window.__ultraUniverAPI;
        (window as any).univerAPI = univerAPI;

        return;
      }

      const { univerAPI, univer } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: locales,
        },
        presets,
        darkMode: resolvedTheme === "dark",
        theme: resolvedTheme === "dark" ? greenTheme : defaultTheme,
      });

      // Lifecycle hook removed: initial formula computing already forced via preset config

      // Create workbook
      univerAPI.createWorkbook({});

      // Attempt official watermark removal once
      try {
        removeUniverWatermark(univerAPI);
      } catch {}

      // EXPOSE univerAPI GLOBALLY for chat component and guard flags
      (window as any).univerAPI = univerAPI;
      window.__ultraUniverAPI = univerAPI;
      window.__ultraUniverInitialized = true;

      // Provide a dispose hook (not used during HMR reuse, but available)
      window.__ultraUniverDispose = () => {
        try {
          univer.dispose?.();
        } catch {}
        window.__ultraUniverInitialized = false;
        window.__ultraUniverAPI = undefined;
        (window as any).univerAPI = undefined;
      };

      // Initialize modern tool system only
      const { setupModernUniverBridge } = await import(
        "../lib/modern-univer-bridge"
      );

      (window as any).__attachUltraDispatcher = () => {
        try {
          setupModernUniverBridge();
        } catch (modernError) {
          console.error(
            "❌ Failed to initialize modern tool system:",
            modernError
          );
          throw modernError; // Don't fall back to legacy - modern system is required
        }
      };

      // Attach dispatcher
      try {
        (window as any).__attachUltraDispatcher?.();
      } catch (e) {
        console.error("❌ Failed to attach modern tool dispatcher:", e);
        throw e;
      }

      // // Run implementation tests in development
      // if (process.env.NODE_ENV === 'development') {
      //   try {
      //     const { testImplementationImprovements } = await import("../lib/test-improvements");
      //     // Run test after everything is initialized
      //     setTimeout(() => {
      //       testImplementationImprovements().catch(console.error);
      //     }, 3000);
      //   } catch (testError) {
      //     console.warn("⚠️ Could not load test improvements:", testError);
      //   }
      // }
    };

    initializeUniver().catch(console.error);

    // Cleanup function to prevent issues during HMR
    return () => {
      try {
        // Restore original console functions
        console.error = originalError;
        console.warn = originalWarn;
      } catch (e) {
        originalWarn("Cleanup warning:", e);
      }
    };
  }, []);

  // Sync Univer dark mode when the app theme changes (including reuse path)
  useEffect(() => {
    try {
      const api: any =
        (window as any)?.__ultraUniverAPI || (window as any)?.univerAPI;
      if (api?.toggleDarkMode) api.toggleDarkMode(resolvedTheme === "dark");
    } catch {}
  }, [resolvedTheme]);

  // No watermark heuristics; rely on official API call once during init

  return (
    <div
      ref={containerRef}
      className="h-full rounded-2xl overflow-hidden m-2 overflow-x-hidden"
    />
  );
}
