// services/chartService.ts
import { UniverService } from "./univerService";
import QuickChart from "quickchart-js";

interface ChartData {
  labels: string[];
  datasets: Array<{
    label: string;
    data: number[];
    backgroundColor?: string | string[];
    borderColor?: string;
  }>;
}

interface ChartResult {
  chartUrl: string;
  dashboardCell: string;
}

export class ChartService {
  private static instance: ChartService;
  private univerService: UniverService;

  private constructor() {
    this.univerService = UniverService.getInstance();
  }

  static getInstance(): ChartService {
    if (!ChartService.instance) {
      ChartService.instance = new ChartService();
    }
    return ChartService.instance;
  }

  async generateChart(
    data: any[],
    xColumn: string,
    yColumn: string,
    chartType: string,
    title: string
  ): Promise<ChartResult> {
    try {
      // Transform data for chart
      const chartData = this.transformDataForChart(
        data,
        xColumn,
        yColumn,
        chartType
      );

      // Generate chart URL using QuickChart
      const chartUrl = await this.generateQuickChartUrl(
        chartData,
        chartType,
        title
      );

      // Insert chart to Dashboard sheet
      const dashboardCell = await this.insertChartToDashboard(chartUrl, title);

      return {
        chartUrl,
        dashboardCell,
      };
    } catch (error) {
      throw new Error(
        `Failed to generate chart: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  private transformDataForChart(
    data: any[],
    xColumn: string,
    yColumn: string,
    chartType: string
  ): ChartData {
    if (chartType === "bar" || chartType === "line") {
      return {
        labels: data.map((row) => String(row[xColumn] || "")),
        datasets: [
          {
            label: yColumn,
            data: data.map((row) => parseFloat(row[yColumn]) || 0),
            backgroundColor: "rgba(54, 162, 235, 0.5)",
            borderColor: "rgba(54, 162, 235, 1)",
          },
        ],
      };
    } else if (chartType === "pie") {
      return {
        labels: data.map((row) => String(row[xColumn] || "")),
        datasets: [
          {
            label: yColumn,
            data: data.map((row) => parseFloat(row[yColumn]) || 0),
            backgroundColor: [
              "rgba(255, 99, 132, 0.5)",
              "rgba(54, 162, 235, 0.5)",
              "rgba(255, 206, 86, 0.5)",
              "rgba(75, 192, 192, 0.5)",
              "rgba(153, 102, 255, 0.5)",
            ],
          },
        ],
      };
    } else if (chartType === "scatter") {
      return {
        labels: data.map((row) => String(row[xColumn] || "")),
        datasets: [
          {
            label: yColumn,
            data: data.map((row) => parseFloat(row[yColumn]) || 0),
            backgroundColor: "rgba(255, 99, 132, 0.5)",
          },
        ],
      };
    }

    throw new Error(`Unsupported chart type: ${chartType}`);
  }

  private async generateQuickChartUrl(
    chartData: ChartData,
    chartType: string,
    title: string
  ): Promise<string> {
    // Create QuickChart instance
    const chart = new QuickChart();

    chart.setConfig({
      type: chartType,
      data: chartData,
      options: {
        responsive: true,
        plugins: {
          title: {
            display: true,
            text: title,
          },
          legend: {
            display: true,
          },
        },
      },
    });

    // Generate PNG URL
    return await chart.getUrl();
  }

  private async insertChartToDashboard(
    chartUrl: string,
    title: string
  ): Promise<string> {
    // For now, just return a placeholder cell
    // In a full implementation, this would create a Dashboard sheet and insert the chart
    console.log(`Chart would be inserted at A1 with title: ${title}`);
    console.log(`Chart URL: ${chartUrl}`);

    return "A1"; // Placeholder
  }
}
