const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const Fi = require("react-icons/fi");
const Fa = require("react-icons/fa");

// ---------- palette (matches manavcapital site) ----------
const INK = "10192B";
const INK_SOFT = "1C2A40";
const INK_LINE = "33415A";
const INK_MUTE = "616B80";
const CANVAS = "F4EFE3";
const CANVAS_DEEP = "E9DFC7";
const PAPER = "FBF8F1";
const GOLD = "B4863D";
const GOLD_DEEP = "8A6529";
const GOLD_SOFT = "D9BC84";
const RULE = "D8CFB6";
const WHITE = "FFFFFF";
const MUTED_DARK = "A9B2C4";

const F_HEAD = "Cambria";
const F_BODY = "Calibri";

// ---------- icon rasterizer ----------
const iconCache = {};
async function iconPng(Comp, colorHex, name) {
  const key = name + colorHex;
  if (iconCache[key]) return iconCache[key];
  // Keep react-icons' own root <svg> (correct native viewBox + currentColor
  // context) instead of rebuilding it — that rebuild was the bug: it forced
  // every icon into a 512 viewBox, shrinking 24x24 (Feather) icons to a dot,
  // and dropped the color styling FontAwesome icons carry on the root tag.
  const svgMarkup = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { size: 512, color: "#" + colorHex })
  );
  const buf = await sharp(Buffer.from(svgMarkup)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  const data = "image/png;base64," + buf.toString("base64");
  iconCache[key] = data;
  return data;
}

function badgeCircle(slide, x, y, d, ringColor) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: "FFFFFF", transparency: 100 }, line: { color: ringColor, width: 1 } });
}

async function main() {
  const pres = new pptxgen();
  pres.defineLayout({ name: "MANAV_WIDE", width: 13.333, height: 7.5 });
  pres.layout = "MANAV_WIDE";
  const W = 13.333, H = 7.5;

  // pre-render icons (dark-bg gold, light-bg gold-deep)
  const ic = {};
  ic.jetGold = await iconPng(Fa.FaFighterJet, GOLD, "jet");
  ic.usersGold = await iconPng(Fi.FiUsers, GOLD, "users");
  ic.homeGold = await iconPng(Fi.FiHome, GOLD, "home");
  ic.hammerGold = await iconPng(Fa.FaHammer, GOLD, "hammer");
  ic.arrowsGold = await iconPng(Fi.FiRefreshCw, GOLD, "arrows");
  ic.chartGold = await iconPng(Fi.FiTrendingUp, GOLD, "chart");
  ic.shieldGoldDeep = await iconPng(Fi.FiShield, GOLD_DEEP, "shieldd");
  ic.hammerGoldDeep = await iconPng(Fa.FaHammer, GOLD_DEEP, "hammerd");
  ic.chartGoldDeep = await iconPng(Fi.FiTrendingUp, GOLD_DEEP, "chartd");
  ic.usersGoldDeep = await iconPng(Fi.FiUsers, GOLD_DEEP, "usersd");
  ic.keyGoldDeep = await iconPng(Fi.FiKey, GOLD_DEEP, "keyd");
  ic.arrowsGoldDeep = await iconPng(Fi.FiRefreshCw, GOLD_DEEP, "arrowsd");
  ic.buildingGoldSoft = await iconPng(Fa.FaBuilding, GOLD_SOFT, "buildings");
  ic.homeGoldSoft = await iconPng(Fi.FiHome, GOLD_SOFT, "homes");
  ic.mailGoldSoft = await iconPng(Fi.FiMail, GOLD_SOFT, "mails");
  ic.logo = "image/png;base64," + fs.readFileSync(path.join(__dirname, "..", "assets", "logo-icon.png")).toString("base64");

  const footerBrand = (slide, dark) => {
    slide.addText("MANAV CAPITAL GROUP", { x: 0.6, y: H - 0.55, w: 4, h: 0.3, fontFace: F_BODY, fontSize: 9, color: dark ? MUTED_DARK : INK_MUTE, charSpacing: 1, margin: 0 });
  };
  const pageNo = (slide, n, dark) => {
    slide.addText(String(n).padStart(2, "0"), { x: W - 1.1, y: H - 0.55, w: 0.6, h: 0.3, align: "right", fontFace: F_BODY, fontSize: 9, color: dark ? MUTED_DARK : INK_MUTE, margin: 0 });
  };
  const kicker = (slide, text, x, y, w, dark) => {
    slide.addText(text.toUpperCase(), { x, y, w, h: 0.35, fontFace: F_BODY, bold: true, fontSize: 11, color: dark ? GOLD : GOLD_DEEP, charSpacing: 2, margin: 0 });
  };
  const placeholderBox = (slide, x, y, w, h, label) => {
    slide.addShape("rect", { x, y, w, h, fill: { color: INK }, line: { color: GOLD, width: 1 } });
    slide.addImage({ data: ic.homeGoldSoft, x: x + w / 2 - 0.35, y: y + h / 2 - 0.75, w: 0.7, h: 0.7 });
    slide.addText(label, { x: x + 0.4, y: y + h / 2 + 0.05, w: w - 0.8, h: 0.6, align: "center", fontFace: F_BODY, fontSize: 10.5, bold: true, color: GOLD_SOFT, charSpacing: 1, margin: 0 });
    slide.addText("PLACEHOLDER — ADD REAL PHOTOGRAPHY", { x: x + 0.4, y: y + h - 0.5, w: w - 0.8, h: 0.3, align: "center", fontFace: F_BODY, fontSize: 8, color: GOLD_SOFT, charSpacing: 1, margin: 0 });
  };

  // ============================================================ SLIDE 1 — TITLE
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    s.addImage({ data: ic.logo, x: W / 2 - 0.5, y: 1.35, w: 1.0, h: 1.0, transparency: 0 });
    s.addText("REAL ESTATE CAPITAL INVESTMENT · INVESTOR OVERVIEW", { x: 0, y: 2.55, w: W, h: 0.4, align: "center", fontFace: F_BODY, bold: true, fontSize: 12, color: GOLD, charSpacing: 2, margin: 0 });
    s.addText("731 N Davis Dr", { x: 0, y: 3.05, w: W, h: 1.1, align: "center", fontFace: F_HEAD, fontSize: 54, color: CANVAS, margin: 0 });
    s.addText("Warner Robins, Georgia", { x: 0, y: 4.05, w: W, h: 0.5, align: "center", fontFace: F_HEAD, italic: true, fontSize: 22, color: GOLD_SOFT, margin: 0 });
    s.addText("A value-add multifamily acquisition minutes from Robins Air Force Base — presented by MANAV Capital Group, the human touch.", { x: W / 2 - 3.6, y: 4.75, w: 7.2, h: 0.8, align: "center", fontFace: F_BODY, fontSize: 13, color: MUTED_DARK, margin: 0 });
    footerBrand(s, true); pageNo(s, 1, true);
  }

  // ============================================================ SLIDE 2 — DISCLOSURE
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    kicker(s, "Disclosure & Forward-Looking Statements", 0.6, 0.55, 8, false);
    s.addText("Please read before proceeding.", { x: 0.6, y: 0.9, w: 10, h: 0.6, fontFace: F_HEAD, fontSize: 28, color: INK, margin: 0 });
    const paras = [
      "This Investor Overview has been prepared solely for, and is being delivered on a confidential basis to, persons considering an investment with MANAV Capital Group. It is for informational purposes only and does not constitute an offer to sell, a solicitation of an offer to buy, or a recommendation to purchase any interest in this or any property. Any offer will be made only pursuant to definitive transaction documents.",
      "This document contains forward-looking statements — including projected returns, occupancy, rent growth, and exit outcomes — identified by words such as “target,” “project,” “estimate,” “expect,” and similar terms. These statements are based on current assumptions and are subject to risks and uncertainties that could cause actual results to differ materially. Past performance of MANAV Capital Group or any other sponsor is not a guarantee of future results.",
      "Market and economic data referenced herein (including Robins Air Force Base economic impact and Warner Robins housing statistics) is drawn from publicly available third-party sources believed to be reliable; MANAV Capital Group has not independently verified this data and makes no representation as to its accuracy or completeness. Comparable sales data in this deck is presented as an illustrative template pending final verification against county records, MLS, or broker-sourced data.",
      "Prospective investors should conduct their own due diligence and consult independent tax, legal, and financial advisors before making any investment decision.",
    ];
    let y = 1.75;
    paras.forEach((p, i) => {
      s.addText(p, { x: 0.6, y, w: 12.1, h: i === 3 ? 0.5 : 1.05, fontFace: F_BODY, fontSize: 11.5, italic: i === 3, color: i === 3 ? INK_MUTE : INK_SOFT, valign: "top", margin: 0, lineSpacingMultiple: 1.25 });
      y += i === 3 ? 0.55 : 1.15;
    });
    footerBrand(s, false); pageNo(s, 2, false);
  }

  // ============================================================ SLIDE 3 — THE OPPORTUNITY
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    kicker(s, "The Opportunity", 0.6, 0.6, 6, false);
    s.addText([{ text: "An undervalued fourplex beside a ", options: {} }, { text: "$4.48B", options: { italic: true, color: GOLD_DEEP } }, { text: " economic engine.", options: {} }], { x: 0.6, y: 1.0, w: 6.9, h: 1.5, fontFace: F_HEAD, fontSize: 30, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
    s.addText("731 N Davis Dr is a 4-unit, 1973-built multifamily property in Warner Robins, Georgia — currently renting well below what a renovated unit commands in this market. The surrounding economy is anchored by Robins Air Force Base, whose workforce has grown 11% in two years and is expected to keep expanding.", { x: 0.6, y: 2.75, w: 6.7, h: 1.3, fontFace: F_BODY, fontSize: 12.5, color: INK_SOFT, margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("MANAV's plan: acquire, reposition, and stabilize the asset to capture the gap between in-place and market rent — then return capital to investors through a sale or refinance within three years.", { x: 0.6, y: 4.05, w: 6.7, h: 1.0, fontFace: F_BODY, fontSize: 12.5, color: INK_SOFT, margin: 0, lineSpacingMultiple: 1.2 });
    ["VALUE-ADD MULTIFAMILY", "3-YEAR HOLD", "DIRECT OWNERSHIP"].forEach((t, i) => {
      s.addShape("roundRect", { x: 0.6 + i * 2.05, y: 5.15, w: 1.95, h: 0.4, rectRadius: 0.2, fill: { color: PAPER }, line: { color: GOLD_DEEP, width: 1 } });
      s.addText(t, { x: 0.6 + i * 2.05, y: 5.15, w: 1.95, h: 0.4, align: "center", valign: "middle", fontFace: F_BODY, bold: true, fontSize: 8.5, color: GOLD_DEEP, charSpacing: 0.5, margin: 0 });
    });
    placeholderBox(s, 8.0, 1.0, 4.7, 5.6, "Property Photo\n731 N Davis Dr, Warner Robins, GA");
    footerBrand(s, false); pageNo(s, 3, false);
  }

  // ============================================================ SLIDE 4 — WHY WARNER ROBINS
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    kicker(s, "Market Overview", 0.6, 0.55, 6, true);
    s.addText("Why Warner Robins is the next Middle Georgia play.", { x: 0.6, y: 0.9, w: 10.5, h: 0.7, fontFace: F_HEAD, fontSize: 26, color: CANVAS, margin: 0 });
    s.addText("Explosive base-driven employment growth, tightening rental demand, and a housing market still catching up.", { x: 0.6, y: 1.55, w: 10.5, h: 0.4, fontFace: F_BODY, fontSize: 12, color: MUTED_DARK, margin: 0 });

    const cards = [
      { icon: ic.jetGold, num: "$4.48B", lbl: "Robins AFB economic impact in 2025 — the region's single largest economic driver." },
      { icon: ic.usersGold, num: "21,000+", lbl: "Base employees — up 11% in two years, roughly a third of all regional jobs." },
      { icon: ic.homeGold, num: "~2,000", lbl: "Additional households needing housing from base workforce growth alone." },
    ];
    cards.forEach((c, i) => {
      const x = 0.6 + i * 4.15;
      s.addShape("rect", { x, y: 2.3, w: 3.9, h: 2.9, fill: { color: INK_SOFT }, line: { color: INK_LINE, width: 1 } });
      s.addImage({ data: c.icon, x: x + 0.35, y: 2.65, w: 0.55, h: 0.55 });
      s.addText(c.num, { x: x + 0.3, y: 3.35, w: 3.3, h: 0.6, fontFace: F_HEAD, fontSize: 28, color: GOLD, margin: 0 });
      s.addText(c.lbl, { x: x + 0.3, y: 4.0, w: 3.3, h: 1.0, fontFace: F_BODY, fontSize: 11, color: MUTED_DARK, margin: 0, lineSpacingMultiple: 1.2 });
    });
    s.addText("Sources: 13WMAZ, “Robins Air Force Base drives $4.48 billion economic impact,” 2026 State of the Base Address; Walton Dean Realty, “Robins AFB & Houston County Housing Market,” 2026.", { x: 0.6, y: 5.5, w: 11.5, h: 0.5, fontFace: F_BODY, fontSize: 9, color: "8592AA", margin: 0 });
    footerBrand(s, true); pageNo(s, 4, true);
  }

  // ============================================================ SLIDE 5 — DEAL AT A GLANCE
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    kicker(s, "The Deal at a Glance", 0.6, 0.55, 6, false);
    s.addText("Snapshot of key facts and targets.", { x: 0.6, y: 0.9, w: 10, h: 0.6, fontFace: F_HEAD, fontSize: 28, color: INK, margin: 0 });

    const cells = [
      ["$1.0M", "Purchase Price", false],
      ["4", "Residential Units", false],
      ["$250K", "Price Per Unit", false],
      ["8.5%", "Target Cash-on-Cash Return", true],
      ["Up to 15%", "Target Tax-Deferred IRR", true],
      ["3 Yrs", "Hold Period", false],
    ];
    const cw = 11.5 / 6;
    cells.forEach((c, i) => {
      const x = 0.6 + i * cw;
      s.addShape("rect", { x, y: 1.85, w: cw - 0.03, h: 1.5, fill: { color: PAPER }, line: { color: RULE, width: 0.75 } });
      s.addText(c[0], { x: x + 0.12, y: 2.0, w: cw - 0.3, h: 0.65, fontFace: F_HEAD, fontSize: c[0].length > 5 ? 19 : 24, color: c[2] ? GOLD_DEEP : INK, margin: 0 });
      s.addText(c[1], { x: x + 0.12, y: 2.68, w: cw - 0.3, h: 0.6, fontFace: F_BODY, fontSize: 9.5, color: INK_MUTE, margin: 0, lineSpacingMultiple: 1.15 });
    });

    s.addText("Well-located value-add fourplex", { x: 0.6, y: 4.0, w: 5.6, h: 0.4, fontFace: F_BODY, bold: true, fontSize: 14, color: INK, margin: 0 });
    s.addText("Built 1973 · 2,756 sq ft · 1.45-acre lot · Houston County, GA. Located near Robins AFB, schools, and town amenities.", { x: 0.6, y: 4.4, w: 5.6, h: 0.9, fontFace: F_BODY, fontSize: 12, color: INK_SOFT, margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("Below-market in-place rents", { x: 6.55, y: 4.0, w: 5.6, h: 0.4, fontFace: F_BODY, bold: true, fontSize: 14, color: INK, margin: 0 });
    s.addText("Current rents run ~$760–$795/unit/mo. Renovation and repositioning are underwritten to close the gap to market.", { x: 6.55, y: 4.4, w: 5.6, h: 0.9, fontFace: F_BODY, fontSize: 12, color: INK_SOFT, margin: 0, lineSpacingMultiple: 1.2 });
    footerBrand(s, false); pageNo(s, 5, false);
  }

  // ============================================================ SLIDE 6 — PROPERTY OVERVIEW
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    placeholderBox(s, 0.6, 1.0, 4.4, 5.6, "Site Photo\n731 N Davis Dr");

    kicker(s, "Property Overview", 5.4, 0.6, 6, false);
    s.addText("By the numbers: a fourplex built for repositioning.", { x: 5.4, y: 0.95, w: 7.3, h: 0.9, fontFace: F_HEAD, fontSize: 22, color: INK, margin: 0 });

    const rows = [
      ["Address", "731 N Davis Dr, Warner Robins, GA 31093"],
      ["County", "Houston County, GA"],
      ["Year Built", "1973"],
      ["Units", "4"],
      ["Avg. Unit Size", "~600 sq ft"],
      ["Total Building Area", "2,756 sq ft"],
      ["Lot Size", "1.45 acres"],
      ["Current Rent Range", "$760 – $795 / unit / mo"],
      ["Purchase Price", "$1,000,000"],
    ];
    s.addTable(
      rows.map(([k, v]) => ([
        { text: k, options: { bold: true, color: INK_MUTE, fontFace: F_BODY, fontSize: 10.5, fill: { color: PAPER } } },
        { text: v, options: { color: INK, fontFace: F_BODY, fontSize: 11.5, fill: { color: PAPER } } },
      ])),
      { x: 5.4, y: 2.0, w: 7.3, colW: [2.4, 4.9], border: { type: "solid", color: RULE, pt: 0.75 }, autoPage: false, rowH: 0.5 }
    );
    footerBrand(s, false); pageNo(s, 6, false);
  }

  // ============================================================ SLIDE 7 — BUSINESS PLAN
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    kicker(s, "Business Plan", 0.6, 0.55, 6, true);
    s.addText("From acquisition to exit: our three-phase plan.", { x: 0.6, y: 0.9, w: 11, h: 0.6, fontFace: F_HEAD, fontSize: 26, color: CANVAS, margin: 0 });

    const phases = [
      { icon: ic.hammerGold, tag: "PHASE 1 · 0–6 MONTHS", title: "Renovate & Stabilize", d: "Unit-by-unit renovation, deferred maintenance addressed, curb appeal improved." },
      { icon: ic.arrowsGold, tag: "PHASE 2 · 6–18 MONTHS", title: "Reposition to Market Rent", d: "Re-lease renovated units at market rates, targeting base-adjacent renter demand." },
      { icon: ic.chartGold, tag: "PHASE 3 · 18–36 MONTHS", title: "Stabilize & Exit", d: "Hold for stabilized cash flow, then sell or refinance to return investor capital." },
    ];
    phases.forEach((p, i) => {
      const x = 0.6 + i * 4.05;
      s.addShape("rect", { x, y: 1.85, w: 3.8, h: 3.4, fill: { color: "1A2740" }, line: { color: INK_LINE, width: 1 } });
      s.addImage({ data: p.icon, x: x + 0.35, y: 2.15, w: 0.5, h: 0.5 });
      s.addText(p.tag, { x: x + 0.3, y: 2.8, w: 3.2, h: 0.3, fontFace: F_BODY, bold: true, fontSize: 9.5, color: GOLD, charSpacing: 1, margin: 0 });
      s.addText(p.title, { x: x + 0.3, y: 3.12, w: 3.2, h: 0.55, fontFace: F_HEAD, fontSize: 17, color: CANVAS, margin: 0 });
      s.addText(p.d, { x: x + 0.3, y: 3.7, w: 3.2, h: 1.4, fontFace: F_BODY, fontSize: 11, color: MUTED_DARK, margin: 0, lineSpacingMultiple: 1.25 });
    });
    footerBrand(s, true); pageNo(s, 7, true);
  }

  // ============================================================ SLIDE 8 — COMPARABLE SALES
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    kicker(s, "Comparable Sales", 0.6, 0.55, 6, false);
    s.addText([{ text: "Small multifamily comps — ", options: {} }, { text: "pending final verification", options: { italic: true, color: GOLD_DEEP } }, { text: ".", options: {} }], { x: 0.6, y: 0.9, w: 11.5, h: 0.6, fontFace: F_HEAD, fontSize: 26, color: INK, margin: 0 });
    s.addText("Publicly available search could not surface verified closed sales for sub-20-unit properties in this submarket — that data sits behind paid CRE platforms (CoStar/Crexi) or the local MLS. The rows below are a template for your broker or MLS pull; one active listing is shown for general market context only.", { x: 0.6, y: 1.55, w: 11.8, h: 0.75, fontFace: F_BODY, fontSize: 11, color: INK_SOFT, margin: 0, lineSpacingMultiple: 1.2 });

    const head = ["Address", "Units", "Sale Date", "Sale Price", "Price / Unit", "Status"].map(t => ({ text: t, options: { bold: true, color: INK, fontFace: F_BODY, fontSize: 10, fill: { color: "EFE7D2" } } }));
    const mkRow = (cells, italic) => cells.map(t => ({ text: t, options: { color: italic ? INK_MUTE : INK, italic: !!italic, fontFace: F_BODY, fontSize: 10.5, fill: { color: PAPER } } }));
    const rows = [
      head,
      mkRow(["[Comp 1 Address]", "[N]", "[Date]", "[$]", "[$/Unit]", "Pending broker data"], true),
      mkRow(["[Comp 2 Address]", "[N]", "[Date]", "[$]", "[$/Unit]", "Pending broker data"], true),
      mkRow(["[Comp 3 Address]", "[N]", "[Date]", "[$]", "[$/Unit]", "Pending broker data"], true),
      mkRow(["Warner Robins, GA 31088", "96", "—", "$6,500,000", "~$67,700", "Active listing (context only)"], false),
    ];
    s.addTable(rows, { x: 0.6, y: 2.45, w: 12.1, colW: [3.5, 1.1, 1.6, 1.9, 1.8, 2.2], border: { type: "solid", color: RULE, pt: 0.75 }, autoPage: false, rowH: 0.42 });

    s.addText("731 N Davis Dr underwrites at $250,000/unit — above the 96-unit reference asset's per-unit ask, reflecting the smaller asset class and renovated-unit repositioning premium. Confirm against true closed comps before finalizing underwriting.", { x: 0.6, y: 4.9, w: 12.1, h: 0.6, fontFace: F_BODY, italic: true, fontSize: 10, color: INK_MUTE, margin: 0, lineSpacingMultiple: 1.2 });
    footerBrand(s, false); pageNo(s, 8, false);
  }

  // ============================================================ SLIDE 9 — TARGET RETURNS
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    kicker(s, "Target Investor Returns", 0.6, 0.55, 6, true);
    s.addText("What this repositioning is underwritten to deliver.", { x: 0.6, y: 0.9, w: 10.5, h: 0.7, fontFace: F_HEAD, fontSize: 26, color: CANVAS, margin: 0 });

    const boxes = [
      { label: "Cash-on-Cash Return", num: "8.5", suffix: "%", d: "Target annual cash distribution once units are repositioned and stabilized." },
      { label: "Tax-Deferred IRR", num: "Up to 15", suffix: "%", d: "Target internal rate of return over the 3-year hold, inclusive of depreciation-driven tax deferral." },
    ];
    boxes.forEach((b, i) => {
      const x = 0.6 + i * 4.55;
      s.addShape("rect", { x, y: 1.95, w: 4.3, h: 2.7, fill: { color: "1A2740" }, line: { color: INK_LINE, width: 1 } });
      s.addText(b.label, { x: x + 0.35, y: 2.2, w: 3.6, h: 0.4, fontFace: F_HEAD, italic: true, fontSize: 15, color: CANVAS, margin: 0 });
      s.addText([{ text: b.num, options: { fontSize: 44 } }, { text: b.suffix, options: { fontSize: 20 } }], { x: x + 0.3, y: 2.65, w: 3.7, h: 0.9, fontFace: F_HEAD, color: GOLD, margin: 0 });
      s.addText(b.d, { x: x + 0.35, y: 3.65, w: 3.6, h: 0.85, fontFace: F_BODY, fontSize: 11, color: MUTED_DARK, margin: 0, lineSpacingMultiple: 1.2 });
    });
    s.addText("Forward-looking target, not a guarantee. Actual returns depend on renovation cost, lease-up pace, market rent growth, and exit execution. See disclosure, slide 2.", { x: 0.6, y: 5.05, w: 10, h: 0.6, fontFace: F_BODY, fontSize: 10, color: "8592AA", margin: 0, lineSpacingMultiple: 1.2 });
    footerBrand(s, true); pageNo(s, 9, true);
  }

  // ============================================================ SLIDE 10 — EXIT STRATEGY
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    kicker(s, "Exit Strategy", 0.6, 0.55, 6, false);
    s.addText("Two paths to return investor capital.", { x: 0.6, y: 0.9, w: 10, h: 0.6, fontFace: F_HEAD, fontSize: 28, color: INK, margin: 0 });

    const opts = [
      { icon: ic.keyGoldDeep, title: "Option 1 — Sale", d: "Sell the stabilized, fully-repositioned asset at the end of the hold period.", num: "45%", lbl: "Target Total ROI" },
      { icon: ic.arrowsGoldDeep, title: "Option 2 — Refinance", d: "Cash-out refinance once stabilized, returning investor capital while MANAV retains the asset.", num: "7%", lbl: "Target Cash-Out Rate" },
    ];
    opts.forEach((o, i) => {
      const x = 0.6 + i * 5.9;
      s.addShape("rect", { x, y: 1.85, w: 5.6, h: 3.3, fill: { color: PAPER }, line: { color: RULE, width: 1 } });
      s.addImage({ data: o.icon, x: x + 0.35, y: 2.15, w: 0.5, h: 0.5 });
      s.addText(o.title, { x: x + 0.35, y: 2.8, w: 4.9, h: 0.45, fontFace: F_HEAD, fontSize: 19, color: INK, margin: 0 });
      s.addText(o.d, { x: x + 0.35, y: 3.28, w: 4.9, h: 0.75, fontFace: F_BODY, fontSize: 11.5, color: INK_SOFT, margin: 0, lineSpacingMultiple: 1.2 });
      s.addShape("line", { x: x + 0.35, y: 4.15, w: 4.9, h: 0, line: { color: RULE, width: 1 } });
      s.addText(o.num, { x: x + 0.35, y: 4.3, w: 1.7, h: 0.6, fontFace: F_HEAD, fontSize: 30, color: GOLD_DEEP, margin: 0 });
      s.addText(o.lbl.toUpperCase(), { x: x + 1.9, y: 4.45, w: 3.3, h: 0.4, fontFace: F_BODY, fontSize: 9.5, color: INK_MUTE, charSpacing: 0.5, valign: "middle", margin: 0 });
    });
    s.addText("Sale ROI and refinance rate are underwriting targets, not guarantees — actual terms depend on market conditions, lender criteria, and appraised value at time of exit.", { x: 0.6, y: 5.45, w: 11.5, h: 0.5, fontFace: F_BODY, italic: true, fontSize: 10, color: INK_MUTE, margin: 0, lineSpacingMultiple: 1.2 });
    footerBrand(s, false); pageNo(s, 10, false);
  }

  // ============================================================ SLIDE 11 — RISK FACTORS
  {
    const s = pres.addSlide();
    s.background = { color: CANVAS };
    kicker(s, "Risk Factors", 0.6, 0.55, 6, false);
    s.addText("Risk is real. So is our plan to manage it.", { x: 0.6, y: 0.9, w: 10.5, h: 0.6, fontFace: F_HEAD, fontSize: 27, color: INK, margin: 0 });

    const head = ["Risk", "Mitigation"].map((t, i) => ({ text: t, options: { bold: true, color: INK, fontFace: F_BODY, fontSize: 11, fill: { color: "EFE7D2" } } }));
    const data = [
      ["Renovation cost or timeline overrun", "Fixed-scope contractor bids, contingency reserve built into the renovation budget, licensed owner-managed crew."],
      ["Slower-than-planned lease-up at market rents", "Conservative underwriting on rent-up pace; direct marketing to base-adjacent renter demand."],
      ["Interest rate movement at refinance", "Refinance target rate stress-tested against current forecasts; sale remains the fallback exit."],
      ["Local market or base-driven demand shift", "Diversified MANAV portfolio across renovation, multifamily, and STR/MTR strategies limits single-asset exposure."],
    ];
    const rows = [head, ...data.map(([a, b]) => ([
      { text: a, options: { bold: true, color: INK, fontFace: F_BODY, fontSize: 11, fill: { color: PAPER } } },
      { text: b, options: { color: INK_SOFT, fontFace: F_BODY, fontSize: 11, fill: { color: PAPER } } },
    ]))];
    s.addTable(rows, { x: 0.6, y: 1.75, w: 12.1, colW: [3.6, 8.5], border: { type: "solid", color: RULE, pt: 0.75 }, autoPage: false, rowH: [0.5, 0.85, 0.85, 0.85, 0.85] });
    footerBrand(s, false); pageNo(s, 11, false);
  }

  // ============================================================ SLIDE 12 — WHY MANAV
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    kicker(s, "Why MANAV", 0.6, 0.55, 6, true);
    s.addText([{ text: "One team. Renovate, lease, manage, ", options: {} }, { text: "reposition", options: { italic: true, color: GOLD } }, { text: ".", options: {} }], { x: 0.6, y: 0.9, w: 10, h: 0.7, fontFace: F_HEAD, fontSize: 26, color: CANVAS, margin: 0 });

    const items = [
      { icon: ic.shieldGoldDeep, t: "Direct ownership", d: "No syndication — investors hold a direct interest in the asset." },
      { icon: ic.hammerGoldDeep, t: "Hands-on execution", d: "Renovation and leasing run by the same team that underwrote the deal." },
      { icon: ic.chartGoldDeep, t: "Transparent underwriting", d: "Every target return is modeled line-by-line, not promised in the abstract." },
      { icon: ic.usersGoldDeep, t: "The human touch", d: "Real people managing real properties — not a call center." },
    ];
    items.forEach((it, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.6 + col * 5.9, y = 2.0 + row * 1.85;
      s.addShape("ellipse", { x, y, w: 0.6, h: 0.6, fill: { color: "1A2740" }, line: { color: GOLD, width: 1 } });
      s.addImage({ data: it.icon, x: x + 0.13, y: y + 0.13, w: 0.34, h: 0.34 });
      s.addText(it.t, { x: x + 0.8, y: y - 0.02, w: 4.7, h: 0.4, fontFace: F_BODY, bold: true, fontSize: 14.5, color: CANVAS, margin: 0 });
      s.addText(it.d, { x: x + 0.8, y: y + 0.38, w: 4.7, h: 0.7, fontFace: F_BODY, fontSize: 11, color: MUTED_DARK, margin: 0, lineSpacingMultiple: 1.2 });
    });
    footerBrand(s, true); pageNo(s, 12, true);
  }

  // ============================================================ SLIDE 13 — CLOSING / CONTACT
  {
    const s = pres.addSlide();
    s.background = { color: INK };
    s.addText("GET IN TOUCH", { x: 0, y: 1.55, w: W, h: 0.35, align: "center", fontFace: F_BODY, bold: true, fontSize: 11, color: GOLD, charSpacing: 2, margin: 0 });
    s.addText([{ text: "Ready to invest with ", options: {} }, { text: "the human touch", options: { italic: true, color: GOLD } }, { text: "?", options: {} }], { x: 0, y: 2.0, w: W, h: 0.8, align: "center", fontFace: F_HEAD, fontSize: 32, color: CANVAS, margin: 0 });
    s.addText("Request full underwriting, renovation budget detail, and current availability for 731 N Davis Dr.", { x: W / 2 - 3.4, y: 2.85, w: 6.8, h: 0.6, align: "center", fontFace: F_BODY, fontSize: 13, color: MUTED_DARK, margin: 0 });
    s.addImage({ data: ic.mailGoldSoft, x: W / 2 - 1.65, y: 3.75, w: 0.28, h: 0.28 });
    s.addText("info@manavcapital.com", { x: W / 2 - 1.3, y: 3.72, w: 3.5, h: 0.35, fontFace: F_BODY, fontSize: 14, color: GOLD_SOFT, margin: 0 });
    s.addText("(555) 010-0143  (placeholder number)", { x: 0, y: 4.2, w: W, h: 0.35, align: "center", fontFace: F_BODY, fontSize: 13, color: GOLD_SOFT, margin: 0 });
    s.addText("This document is a sample investor overview. Figures marked as targets are forward-looking and not guaranteed; comparable sales are a template pending verification. See full disclosure on slide 2.", { x: W / 2 - 4.2, y: 5.3, w: 8.4, h: 0.7, align: "center", fontFace: F_BODY, fontSize: 9.5, color: "7A8399", margin: 0, lineSpacingMultiple: 1.2 });
    footerBrand(s, true); pageNo(s, 13, true);
  }

  const outPath = path.join(__dirname, "MANAV-731-N-Davis-Dr-Investor-Deck.pptx");
  await pres.writeFile({ fileName: outPath });
  console.log("Wrote " + outPath);
}

main().catch(e => { console.error(e); process.exit(1); });
