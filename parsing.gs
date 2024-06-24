/* The "Intern" age is an implementation detail. */
ages = ["Adult", "Child", "Intern"]
/* These genders are not representative of the totality, but because I'm
 * programming to a spec these are the choices. */
genders = ["Male", "Female", "Unknown", "Other"]
/* This sounds wrong but it's the data we need to output. */
ethnicities = ["Hispanic", "Not Hispanic"]
/* Sorry in advance. */
races = [
  "American Indian/Alaskan Native",
  "Asian",
  "Black/African American",
  "Native Hawaiian/Other Pacific Islander",
  "White",
  "More than one race",
  "Unknown",
]

/* These are for brevity -- they're as they appear on the form.
 * The form should be at ./demographic-information-form.pdf. */
age_shortcodes = ["A", "C", "I"]
gender_shortcodes = ["M", "F", "U", "O"]
ethnicity_shortcodes = ["H", "NH"]
race_shortcodes = ["AI/AN", "AS", "B/AA", "NH/OPI", "WH", "1+", "UNK"]

function parseDemographicsLine(line) {
  /* Take a row of the Demographics Data spreadsheet, then parse it into a dictionary. */
  site_name = line[0]
  count = Number(line[1])
  age = ages[age_shortcodes.indexOf(line[2])]
  gender = genders[gender_shortcodes.indexOf(line[3])]
  ethnicity = ethnicities[ethnicity_shortcodes.indexOf(line[4])]
  race = races[race_shortcodes.indexOf(line[5])]

  return Array(count).fill({
    site: site_name,
    age: age,
    gender: gender,
    ethnicity: ethnicity,
    race: race,
  })
}

function parseDemographicsLineTest() {
  console.log(parseDemographicsLine([
    "Testing site",
    5,
    "A",
    "M",
    "NH",
    "B/AA"
  ]))
}