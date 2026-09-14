// Hardcoded snapshot of the shared workbook's Projects and Credentials tabs
// (Hardware Tracker Data.xlsx), pulled 2026-09-14. Phase 1 only — once
// Power Automate is wired up, this file gets replaced by a live read of
// those same tabs instead of a static snapshot.

const PROJECTS = [
  { name: "testingproj123", archived: false },
  { name: "MM Wave Testing", archived: false },
  { name: "Pegasus", archived: false },
  { name: "Twilight", archived: false },
  { name: "Tahini", archived: true },
];

const CREDENTIALS = [
  { name: "Blake", login: "blake-DCHW", role: "Admin" },
  { name: "Brian", login: "brian-DCHW", role: "Admin" },
  { name: "David", login: "david-DCHW", role: "Admin" },
  { name: "Drew", login: "drew-DCHW", role: "Admin" },
  { name: "Lesther", login: "lesther-DCHW", role: "Admin" },
  { name: "Melissa", login: "melissa-DCHW", role: "Admin" },
  { name: "Riley", login: "riley-DCHW", role: "Admin" },
  { name: "Sandeep", login: "sandeep-DCHW", role: "Admin" },
  { name: "Sander", login: "sander-DCHW", role: "Admin" },
  { name: "Centific Admin", login: "centificadmin-DCHW", role: "Admin" },
];
