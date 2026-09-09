import * as React from "react";
import styles from "./OzoneDay.module.scss";
import type { IOzoneDayProps } from "./IOzoneDayProps";
import { SPHttpClient } from "@microsoft/sp-http";

type Q = { q: string; o: string[]; a: number };
type S = {
  view: "start" | "quiz" | "review" | "done";
  n: number;
  pick?: number;
  answers: { [key: number]: number };
  skipped: number[];
  left: number;
  saveStatus?: "saving" | "saved" | "failed";
  saveError?: string;
  showTwoMinuteWarning?: boolean;
};
const rawQuestions: [string, string[], number][] = [
  [
    "When is the International Day for the Preservation of the Ozone Layer observed every year?",
    ["5 June", "16 September", "22 April", "23 March"],
    1,
  ],
  [
    "What is the official theme of World Ozone Day 2026?",
    [
      "From Science to Global Action",
      "Montreal Protocol: Advancing Climate Action",
      "Global action for a cooler planet",
      "Ozone Heroes: Protecting Our Future",
    ],
    1,
  ],
  [
    "Which layer of Earth's atmosphere contains most of the protective ozone layer?",
    ["Troposphere", "Stratosphere", "Mesosphere", "Thermosphere"],
    1,
  ],
  [
    "What is the approximate wavelength range of UV-B radiation?",
    ["100–200 nm", "200–280 nm", "280–315 nm", "315–400 nm"],
    2,
  ],
  [
    "Which group of chemicals is historically associated with major damage to the ozone layer?",
    [
      "Chlorofluorocarbons (CFCs)",
      "Oxygen and nitrogen",
      "Water vapour",
      "Carbonates",
    ],
    0,
  ],
  [
    "What is the main purpose of the Montreal Protocol?",
    [
      "To control substances that deplete the ozone layer",
      "To regulate international shipping routes",
      "To establish global weather forecasts",
      "To regulate global fisheries",
    ],
    0,
  ],
  [
    "In which year was the Montreal Protocol signed?",
    ["1972", "1987", "1995", "2016"],
    1,
  ],
  [
    "What does the Kigali Amendment primarily address?",
    [
      "The phase-down of certain HFCs",
      "A global ban on refrigeration",
      "A ban on solar energy",
      "A global ban on oxygen production",
    ],
    0,
  ],
  [
    "What is the “ozone hole” technically referring to?",
    [
      "A physical hole in the atmosphere",
      "A region of severely depleted stratospheric ozone",
      "A gap between atmospheric layers",
      "A region without oxygen",
    ],
    1,
  ],
  [
    "The Kigali Amendment was adopted in which year?",
    ["2006", "2010", "2016", "2020"],
    2,
  ],
  [
    "Which type of everyday equipment is most directly connected with sustainable cooling?",
    [
      "Refrigerators and air conditioners",
      "Bicycles only",
      "Bookshelves only",
      "Manual umbrellas only",
    ],
    0,
  ],
  [
    "How can energy-efficient cooling help the environment?",
    [
      "It can reduce electricity demand and associated emissions",
      "It eliminates the need for electricity everywhere",
      "It increases energy consumption",
      "It has no connection with climate impacts",
    ],
    0,
  ],
  [
    "What is the relationship between ozone protection and climate action?",
    [
      "They can provide benefits for both the ozone layer and climate",
      "They are completely unrelated",
      "They increase cooling demand",
      "They reduce rainfall",
    ],
    0,
  ],
  [
    "Why is protecting the ozone layer important for human health?",
    [
      "It helps reduce exposure to harmful ultraviolet radiation",
      "It prevents all infectious diseases",
      "It eliminates all air pollution",
      "It prevents every skin disease",
    ],
    0,
  ],
  [
    "Which statement about ozone is most accurate?",
    [
      "Stratospheric ozone protects life, while ground-level ozone can be an air pollutant",
      "All ozone is beneficial",
      "All ozone is harmful",
      "Ozone exists only near Earth’s surface",
    ],
    0,
  ],
  [
    "What is the Vienna Convention mainly concerned with?",
    [
      "Protecting the ozone layer through international cooperation and scientific research",
      "Regulating international air travel",
      "Controlling ocean pollution",
      "Managing international trade",
    ],
    0,
  ],
  [
    "Which is an ozone-depleting substance (ODS)?",
    ["CFC-12", "Carbon dioxide", "Oxygen", "Nitrogen"],
    0,
  ],
  [
    "What does “phase-down” of HFCs mean?",
    [
      "Gradually reducing the production and consumption of HFCs",
      "Banning refrigeration immediately",
      "Increasing HFC use",
      "Replacing oxygen with HFCs",
    ],
    0,
  ],
  [
    "How can consumers contribute to sustainable cooling?",
    [
      "Choosing energy-efficient equipment and ensuring proper servicing and disposal",
      "Keeping ACs running continuously",
      "Using older inefficient equipment",
      "Releasing refrigerants",
    ],
    0,
  ],
  [
    "Why is sustainable cooling important?",
    [
      "It improves energy efficiency, reduces climate impacts and supports ozone protection",
      "It increases ozone-depleting substances",
      "It ends refrigeration worldwide",
      "It increases electricity consumption",
    ],
    0,
  ],
];

const data: Q[] = rawQuestions.map(([q, o, a]) => ({ q, o, a }));

export default class OzoneDay extends React.Component<IOzoneDayProps, S> {
  private timer?: number;
  private saving: boolean = false;
  public state: S = {
    view: "start",
    n: 0,
    answers: {},
    skipped: [],
    left: 600,
  };
  public componentWillUnmount(): void {
    if (this.timer) window.clearInterval(this.timer);
  }
  private start = (): void => {
    this.setState({ view: "quiz", left: 600, showTwoMinuteWarning: false });
    this.timer = window.setInterval(
      () =>
        this.state.left <= 1
          ? this.done()
          : this.setState((previous) => ({
              left: previous.left - 1,
              showTwoMinuteWarning:
                previous.left === 121 || previous.showTwoMinuteWarning,
            })),
      1000
    );
  };
  private done = (): void => {
    if (this.timer) window.clearInterval(this.timer);
    this.setState({ view: "done" }, this.saveResult);
  };
  private next = (
    answers: { [key: number]: number },
    skipped: number[]
  ): void => {
    const n = this.state.n + 1;
    if (n < 20) this.setState({ n, answers, skipped, pick: undefined });
    else if (skipped.length)
      this.setState({ view: "review", answers, skipped, pick: undefined });
    else {
      if (this.timer) window.clearInterval(this.timer);
      this.setState({ answers, skipped, view: "done" }, this.saveResult);
    }
  };
  private submit = (): void => {
    if (this.state.pick === undefined) return;
    const a = { ...this.state.answers, [this.state.n]: this.state.pick };
    this.next(
      a,
      this.state.skipped.filter((x) => x !== this.state.n)
    );
  };
  private skip = (): void => {
    const s =
      this.state.skipped.indexOf(this.state.n) < 0
        ? [...this.state.skipped, this.state.n]
        : this.state.skipped;
    this.next(this.state.answers, s);
  };
  private clock = (): string =>
    `${Math.floor(this.state.left / 60)}:${
      this.state.left % 60 < 10 ? "0" : ""
    }${this.state.left % 60}`;
  private saveResult = (): void => {
    if (this.saving) return;
    this.saving = true;
    const score = Object.keys(this.state.answers).filter(
      (k) => this.state.answers[+k] === data[+k].a
    ).length;
    const elapsed = 600 - this.state.left;
    const payload = {
      Title: `${this.props.userDisplayName} - Ozone Day Quiz`,
      EmployeeName: this.props.userDisplayName,
      EmployeeEmail: this.props.userEmail,
      Score: score,
      TimeTaken: `${Math.floor(elapsed / 60)}:${elapsed % 60 < 10 ? "0" : ""}${
        elapsed % 60
      }`,
      QuizDate: new Date().toISOString(),
    };
    this.setState({ saveStatus: "saving", saveError: undefined });
    this.props.spHttpClient
      .post(
        `${this.props.listWebUrl}/_api/web/lists/getbytitle('OzoneDayQuestion')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
          },
          body: JSON.stringify(payload),
        }
      )
      .then(async (response) => {
        if (response.ok) {
          this.setState({ saveStatus: "saved" });
          return;
        }
        const detail = await response.text();
        this.saving = false;
        this.setState({
          saveStatus: "failed",
          saveError: `SharePoint returned ${response.status}${
            detail ? `: ${detail.substring(0, 180)}` : ""
          }`,
        });
      })
      .catch((error) => {
        this.saving = false;
        this.setState({
          saveStatus: "failed",
          saveError:
            error instanceof Error ? error.message : "Network request failed.",
        });
      });
  };
  public render(): React.ReactElement<IOzoneDayProps> {
    const s = this.state;
    if (s.view === "start")
      return (
        <section className={styles.ozoneDay}>
          <div className={styles.hero}>
            <small>MY PI · ENVIRONMENTAL AWARENESS</small>
            <div className={styles.o3}>
              O<sub>3</sub>
            </div>
            <h1>
              International <em>Ozone Day</em> Quiz
            </h1>
            <b>16 SEPTEMBER 2026</b>
            <p>
              Hello, {this.props.userDisplayName}. Test your knowledge and help
              protect our shared atmosphere.
            </p>
            <div className={styles.stats}>
              <span>
                <strong>20</strong>questions
              </span>
              <span>
                <strong>10</strong>minutes
              </span>
              <span>
                <strong>1</strong>mark each
              </span>
              <span>
                <strong>0</strong>negative marks
              </span>
            </div>
            <i>Open to PN01 (Panoli) employees</i>
            <button onClick={this.start}>Start quiz →</button>
            <small>Once submitted, an answer cannot be changed.</small>
          </div>
        </section>
      );
    if (s.view === "review")
      return (
        <section className={styles.ozoneDay}>
          <div className={styles.hero}>
            <small>ALMOST THERE</small>
            <h1>Complete unanswered questions</h1>
            <p>
              Submitted answers are locked. Complete every remaining question to
              finish.
            </p>
            <div className={styles.pending}>
              {s.skipped.map((n) => (
                <button
                  key={n}
                  onClick={() =>
                    this.setState({ view: "quiz", n, pick: undefined })
                  }
                >
                  Question {n + 1}
                  <small>Unanswered</small>
                </button>
              ))}
            </div>
            <b>◷ {this.clock()} remaining</b>
          </div>
        </section>
      );
    if (s.view === "done") {
      const score = Object.keys(s.answers).filter(
        (k) => s.answers[+k] === data[+k].a
      ).length;
      const elapsed = 600 - s.left;
      return (
        <section className={styles.ozoneDay}>
          <div className={styles.hero}>
            <div className={styles.tick}>✓</div>
            <small>QUIZ COMPLETED</small>
            <h1>Thank you for taking action.</h1>
            <p>
              {s.saveError
                ? "Your result could not be saved. Please contact the quiz administrator."
                : "Your result has been recorded for the International Ozone Day Quiz."}
            </p>
            <div className={styles.result}>
              <span>
                Your score
                <strong>
                  {score}
                  <small>/ 20</small>
                </strong>
              </span>
              <span>
                Time taken
                <strong>
                  {Math.floor(elapsed / 60)}:{elapsed % 60 < 10 ? "0" : ""}
                  {elapsed % 60}
                </strong>
              </span>
            </div>
            <i>
              The leaderboard ranks participants by score, then by shortest
              completion time.
            </i>
          </div>
        </section>
      );
    }
    const q = data[s.n];
    return (
      <section className={styles.ozoneDay}>
        {s.showTwoMinuteWarning && (
          <div className={styles.warningOverlay}>
            <div className={styles.warningDialog}>
              <strong>2 minutes remaining</strong>
              <p>You have only 2 minutes left to complete the quiz.</p>
              <button
                onClick={() => this.setState({ showTwoMinuteWarning: false })}
              >
                Continue quiz
              </button>
            </div>
          </div>
        )}
        <header>
          <div>
            <small>INTERNATIONAL OZONE DAY</small>
            <h3>
              Question {s.n + 1} <i>/ 20</i>
            </h3>
          </div>
          <b className={s.left < 60 ? styles.urgent : ""}>◷ {this.clock()}</b>
        </header>
        <div className={styles.progress}>
          <span style={{ width: `${Object.keys(s.answers).length * 5}%` }} />
        </div>
        <main>
          <small>
            QUESTION {s.n < 9 ? "0" : ""}
            {s.n + 1}
          </small>
          <h2>{q.q}</h2>
          {q.o.map((o, i) => (
            <button
              className={s.pick === i ? styles.selected : ""}
              key={o}
              onClick={() => this.setState({ pick: i })}
            >
              <span>{String.fromCharCode(65 + i)}</span>
              {o}
            </button>
          ))}
        </main>
        <footer>
          <button onClick={this.skip}>Skip for now</button>
          <button disabled={s.pick === undefined} onClick={this.submit}>
            Submit answer →
          </button>
        </footer>
      </section>
    );
  }
}
