"use client";

import { useEffect, useRef } from "react";
import { useTheme } from "next-themes";
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-drawing/lib/index.css";
import "@univerjs/preset-sheets-advanced/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
// Enable formula facade API (VLOOKUP/INDEX/MATCH and 500+ functions)
import "@univerjs/sheets-formula/facade";

// Global tool execution handler
declare global {
  interface Window {
    executeUniverTool: (toolName: string, params?: any) => Promise<any>;
    lastSumCell?: string; // Track the last sum cell created
    __ultraUniverInitialized?: boolean;
    __ultraUniverAPI?: any;
    __ultraUniverLifecycleHookAdded?: boolean;
    __ultraUniverDispose?: () => void;
    __univerWatermarkObserver?: MutationObserver;
  }
}

export function Univer() {
  const containerRef = useRef<HTMLDivElement>(null!);
  const { resolvedTheme } = useTheme();

  useEffect(() => {
    // Dynamic imports to avoid SSR issues
    const initializeUniver = async () => {
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
      const { LifecycleStages } = await import("@univerjs/core");

      // Try to import filter preset, fallback if it fails
      let filterPreset = null;
      let filterLocales = null;
      try {
        const filterModule = await import("@univerjs/preset-sheets-filter");
        const filterLocaleModule = await import(
          "@univerjs/preset-sheets-filter/locales/en-US"
        );
        filterPreset = filterModule.UniverSheetsFilterPreset;
        filterLocales = filterLocaleModule.default;
        console.log("✅ Filter preset loaded successfully");
      } catch (error) {
        console.warn(
          "⚠️ Filter preset not available, continuing without filters:",
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
      if (TablePreset) {
        presets.push(TablePreset());
      }

      let locales = mergeLocales(sheetsCoreEnUS);
      try {
        locales = filterLocales
          ? mergeLocales(
              sheetsCoreEnUS,
              sheetsDrawingEnUS,
              sheetsAdvancedEnUS,
              sheetsCFEnUS,
              filterLocales,
              tableLocales || {}
            )
          : mergeLocales(
              sheetsCoreEnUS,
              sheetsDrawingEnUS,
              sheetsAdvancedEnUS,
              sheetsCFEnUS,
              tableLocales || {}
            );
      } catch (e) {
        console.warn(
          "⚠️ Failed to merge locales for chart presets; using core locales only.",
          e
        );
        locales = filterLocales
          ? mergeLocales(sheetsCoreEnUS, filterLocales, tableLocales || {})
          : mergeLocales(sheetsCoreEnUS, tableLocales || {});
      }

      // Helper: remove Univer watermark across API variants
      const removeUniverWatermark = (api: any) => {
        try {
          // Try multiple API methods for watermark removal
          if (typeof api?.removeWatermark === "function") api.removeWatermark();
          if (typeof api?.deleteWatermark === "function") api.deleteWatermark();
          if (typeof api?.hideWatermark === "function") api.hideWatermark();

          // CSS-based watermark removal as fallback
          const removeWatermarkCSS = () => {
            const watermarkSelectors = [
              '[class*="watermark"]',
              '[class*="univer-watermark"]',
              '[class*="logo"]',
              "[data-watermark]",
              ".univer-footer",
              ".univer-brand",
              'div[style*="watermark"]',
              'div[style*="univer"]',
            ];

            watermarkSelectors.forEach((selector) => {
              const elements = document.querySelectorAll(selector);
              elements.forEach((el) => {
                if (el instanceof HTMLElement) {
                  el.style.display = "none";
                  el.remove();
                }
              });
            });
          };

          // Apply CSS removal immediately and on DOM changes
          removeWatermarkCSS();

          // Set up observer for dynamic watermark removal
          if (
            typeof window !== "undefined" &&
            !window.__univerWatermarkObserver
          ) {
            const observer = new MutationObserver(() => {
              removeWatermarkCSS();
            });

            observer.observe(document.body, {
              childList: true,
              subtree: true,
              attributes: true,
              attributeFilter: ["class", "style"],
            });

            window.__univerWatermarkObserver = observer;
          }
        } catch (error) {
          console.warn("Watermark removal failed:", error);
        }
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

        console.log(
          "🎯 Univer component: Reusing existing univerAPI (guarded re-init)"
        );
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

      // Remove Univer watermark if API available
      try {
        removeUniverWatermark(univerAPI);
      } catch {}

      // Force initial formula computing at lifecycle Starting for future workbooks
      try {
        if (!window.__ultraUniverLifecycleHookAdded) {
          univerAPI.addEvent(
            univerAPI.Event.LifeCycleChanged,
            ({ stage }: { stage: any }) => {
              if (stage === LifecycleStages.Starting) {
                const formula = univerAPI.getFormula();
                formula.setInitialFormulaComputing(CalculationMode.FORCED);
              }
              // Re-assert watermark removal on lifecycle transitions
              try {
                removeUniverWatermark(univerAPI);
              } catch {}
            }
          );
          window.__ultraUniverLifecycleHookAdded = true;
        }
        // Sync dark mode on mount and when theme changes
        try {
          if (typeof univerAPI.toggleDarkMode === "function") {
            univerAPI.toggleDarkMode(resolvedTheme === "dark");
          }
        } catch {}
      } catch {}

      // Create workbook
      univerAPI.createWorkbook({});

      // Ensure watermark is removed after workbook creation
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
          console.log("🚀 Modern tool system initialized");
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

      console.log("🎯 Univer component: Initialization complete");

      // Additional persistent watermark removal
      setTimeout(() => removeUniverWatermark(univerAPI), 1000);
      setTimeout(() => removeUniverWatermark(univerAPI), 3000);
      setTimeout(() => removeUniverWatermark(univerAPI), 5000);
    };

    initializeUniver().catch(console.error);
  }, []);

  // Sync Univer dark mode when the app theme changes (including reuse path)
  useEffect(() => {
    try {
      const api: any =
        (window as any)?.__ultraUniverAPI || (window as any)?.univerAPI;
      if (api?.toggleDarkMode) api.toggleDarkMode(resolvedTheme === "dark");
    } catch {}
  }, [resolvedTheme]);

  // Persistent watermark removal effect
  useEffect(() => {
    const removeWatermarks = () => {
      try {
        const api: any =
          (window as any)?.__ultraUniverAPI || (window as any)?.univerAPI;
        if (api) {
          // Try API methods
          if (typeof api?.removeWatermark === "function") api.removeWatermark();
          if (typeof api?.deleteWatermark === "function") api.deleteWatermark();
          if (typeof api?.hideWatermark === "function") api.hideWatermark();
        }

        // CSS-based removal
        const watermarkSelectors = [
          '[class*="watermark"]',
          '[class*="univer-watermark"]',
          '[class*="logo"]',
          "[data-watermark]",
          ".univer-footer",
          ".univer-brand",
          'div[style*="watermark"]',
          'div[style*="univer"]',
        ];

        watermarkSelectors.forEach((selector) => {
          const elements = document.querySelectorAll(selector);
          elements.forEach((el) => {
            if (el instanceof HTMLElement) {
              el.style.display = "none";
              el.remove();
            }
          });
        });
      } catch {}
    };

    // Run immediately and set up interval
    removeWatermarks();
    const interval = setInterval(removeWatermarks, 2000);

    return () => clearInterval(interval);
  }, []);

  return (
    <div ref={containerRef} className="h-full rounded-2xl overflow-hidden" />
  );
}
