import Link from "next/link";
import Image from "next/image";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import {
  CheckCircle2,
  Sparkles,
  Columns3,
  MoveRight,
  Bot,
  Layers,
  Workflow,
  LineChart,
} from "lucide-react";

export const metadata = {
  title: "UltraSheets – AI for people who live in spreadsheets",
  description:
    "Ask in plain English. UltraSheets understands your sheet and does the work for you – add columns, move data, build pivots and charts, all safely and visibly.",
};

export default function LandingPage() {
  return (
    <main className="min-h-screen bg-gradient-to-b from-background to-muted/30">
      {/* Hero */}
      <section className="relative px-6 sm:px-10 py-16 sm:py-24 max-w-6xl mx-auto">
        <div className="absolute inset-0 -z-10 bg-[radial-gradient(ellipse_at_top,rgba(99,102,241,0.15),transparent_60%)]" />
        <div className="flex items-center gap-2 mb-6">
          <Badge variant="secondary" className="h-6">
            Designed for Excel-heavy teams
          </Badge>
          <Badge variant="secondary" className="h-6">
            Forget googling formulas
          </Badge>
          <Badge variant="secondary" className="h-6">
            No more #REF!, #ERROR!, or circular loops
          </Badge>
        </div>
        <h1 className="text-4xl sm:text-5xl md:text-6xl font-semibold tracking-tight leading-tight">
          The easiest way to get work done in spreadsheets
        </h1>
        <p className="mt-4 text-muted-foreground text-base sm:text-lg max-w-2xl">
          UltraSheets turns everyday language into precise spreadsheet actions.
          Insert smart columns, move ranges across sheets, build pivots and
          charts – fast, reliable, and explainable.
        </p>
        <div className="mt-8 flex flex-wrap items-center gap-3">
          <Button asChild size="lg" className="rounded-full">
            <Link href="/">Open your sheet</Link>
          </Button>
          <Button asChild variant="outline" size="lg" className="rounded-full">
            <Link href="#features">See what it can do</Link>
          </Button>
          <div className="flex items-center gap-2 text-xs text-muted-foreground">
            <CheckCircle2 className="h-4 w-4" /> No installs. Works in your
            browser.
          </div>
        </div>
        <div className="mt-10 rounded-2xl border bg-card p-3 shadow-sm">
          <div className="flex items-center justify-between px-3 py-2 border-b">
            <div className="flex items-center gap-2 text-xs text-muted-foreground">
              <Bot className="h-4 w-4" /> Live AI Assistant
            </div>
            <div className="flex items-center gap-2 text-xs">
              <Badge variant="secondary">Status: Ready</Badge>
            </div>
          </div>
          <div className="p-4 grid sm:grid-cols-2 gap-4">
            <div className="rounded-lg bg-muted/40 p-4">
              <p className="text-sm text-muted-foreground">You</p>
              <p className="text-sm mt-1">
                Add a column with Unit Price and calculate Total = Quantity ×
                Unit Price
              </p>
            </div>
            <div className="rounded-lg bg-muted/20 p-4">
              <p className="text-sm text-muted-foreground">UltraSheets</p>
              <ul className="text-sm mt-1 space-y-1">
                <li className="flex items-center gap-2">
                  <Sparkles className="h-4 w-4" /> Understands the active table
                  and its headers
                </li>
                <li className="flex items-center gap-2">
                  <Columns3 className="h-4 w-4" /> Inserts a new column right
                  after your data
                </li>
                <li className="flex items-center gap-2">
                  <Workflow className="h-4 w-4" /> Fills formulas down,
                  recalculates safely
                </li>
              </ul>
            </div>
          </div>
        </div>
      </section>

      {/* Features */}
      <section
        id="features"
        className="px-6 sm:px-10 py-12 sm:py-16 max-w-6xl mx-auto"
      >
        <div className="grid md:grid-cols-3 gap-4 sm:gap-6">
          <FeatureCard
            icon={<Columns3 className="h-5 w-5" />}
            title="Smart column placement"
            desc="Adds new columns exactly where they belong – immediately after your data – with formulas filled down correctly."
          />
          <FeatureCard
            icon={<MoveRight className="h-5 w-5" />}
            title="Move anything, anywhere"
            desc="Cut/paste ranges across sheets. If a destination sheet doesn’t exist, UltraSheets creates it for you."
          />
          <FeatureCard
            icon={<Layers className="h-5 w-5" />}
            title="Understands your sheet"
            desc="Uses live context to see tables, headers, and empty space. No cryptic A1 guessing."
          />
          <FeatureCard
            icon={<LineChart className="h-5 w-5" />}
            title="Instant pivots and charts"
            desc="Group, summarize, and visualize in a click – placed neatly beside your data."
          />
          <FeatureCard
            icon={<Sparkles className="h-5 w-5" />}
            title="Natural language actions"
            desc="Ask in plain English: “sum by category”, “format totals as USD”, “switch to March sheet”."
          />
          <FeatureCard
            icon={<Bot className="h-5 w-5" />}
            title="Clear, live feedback"
            desc="Every tool call is visible with success/error badges, so you stay in control."
          />
        </div>
      </section>

      {/* How it works */}
      <section className="px-6 sm:px-10 py-12 sm:py-16 max-w-6xl mx-auto">
        <h2 className="text-2xl sm:text-3xl font-semibold tracking-tight">
          How it works
        </h2>
        <ol className="mt-6 grid sm:grid-cols-3 gap-4 sm:gap-6 text-sm">
          <Step
            n={1}
            title="Explain in plain English"
            text="Type what you want: “add a Unit Price column and totals in USD”."
          />
          <Step
            n={2}
            title="UltraSheets understands your context"
            text="It reads the live sheet – tables, headers, selection – and decides exact ranges."
          />
          <Step
            n={3}
            title="You watch it happen"
            text="Clear badges show actions running, succeeding, or failing, with zero guesswork."
          />
        </ol>
      </section>

      {/* CTA */}
      <section className="px-6 sm:px-10 py-16 max-w-6xl mx-auto">
        <div className="rounded-2xl border bg-card p-8 sm:p-10 text-center">
          <h3 className="text-2xl sm:text-3xl font-semibold">
            Get confident with spreadsheets again
          </h3>
          <p className="mt-2 text-muted-foreground max-w-2xl mx-auto">
            Whether you’re an analyst, a finance lead, or just
            spreadsheet‑curious – UltraSheets makes complex tasks simple, safe,
            and fast.
          </p>
          <div className="mt-6 flex items-center justify-center gap-3">
            <Button asChild size="lg" className="rounded-full">
              <Link href="/">Start using UltraSheets</Link>
            </Button>
            <Button
              asChild
              variant="outline"
              size="lg"
              className="rounded-full"
            >
              <Link href="#features">Explore features</Link>
            </Button>
          </div>
          <div className="mt-4 flex items-center justify-center gap-2 text-xs text-muted-foreground">
            <CheckCircle2 className="h-4 w-4" /> No vendor lock‑in. Your data
            stays in your sheet.
          </div>
        </div>
      </section>

      <footer className="px-6 sm:px-10 py-10 max-w-6xl mx-auto text-xs text-muted-foreground flex items-center justify-between">
        <div className="flex items-center gap-2">
          <Image src="/logo.png" alt="UltraSheets" width={20} height={20} />
          <span>UltraSheets</span>
        </div>
        <div className="flex items-center gap-4">
          <Link className="hover:underline" href="/">
            Open app
          </Link>
          <Link className="hover:underline" href="#features">
            Features
          </Link>
        </div>
      </footer>
    </main>
  );
}

function FeatureCard({
  icon,
  title,
  desc,
}: {
  icon: React.ReactNode;
  title: string;
  desc: string;
}) {
  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm">
      <div className="flex items-center gap-2 text-sm font-medium">
        {icon}
        {title}
      </div>
      <p className="mt-2 text-sm text-muted-foreground">{desc}</p>
    </div>
  );
}

function Step({ n, title, text }: { n: number; title: string; text: string }) {
  return (
    <li className="rounded-xl border bg-card p-5">
      <div className="text-xs text-muted-foreground">Step {n}</div>
      <div className="mt-1 font-medium">{title}</div>
      <p className="mt-1 text-muted-foreground">{text}</p>
    </li>
  );
}
