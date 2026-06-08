import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import ChartDataLabels from 'chartjs-plugin-datalabels';

// Unregister ChartDataLabels from auto-applying to all charts
Chart.unregister(ChartDataLabels);


window.MAPPING_CONFIG = {
  // localStorage key for admin overrides (per-source manual mappings)
  storageKey: "hr_mapping_overrides",

  // Canonical fields. Each entry: { aliases, type, validator?, critical? }
  fields: {
    // ── IDENTITY (merge keys) ──
    "NIK":                       {type:"string", critical:true, aliases:["nik","employee id","emp id","empid","employee number","emp no","empno","staff id","staff number","badge id","badge no","personnel number","pernr","pers no","sap id","user id","user_id"]},
    "Employee Name":             {type:"string", critical:true, aliases:["employee name","emp name","name","full name","employee full name","incumbent","incumbent name","staff name","user name","user_name"]},
    "Email":                     {type:"string", aliases:["email","email address","e-mail","mail"]},

    // ── ORG / POSITION ──
    "Position ID":               {type:"string", aliases:["position id","pos id","position code","job code","job id"]},
    "Position Name":             {type:"string", aliases:["position name","position title","job title","title","position","role","role title","designation","position_name"]},
    "Position Status":           {type:"string", aliases:["position status","status","user status"]},
    "Current Level":             {type:"string", aliases:["current level","level","job level","grade","band","layer","position level","position level / grade","position level grade"]},
    "C-Level":                   {type:"string", aliases:["c-level","c level","clevel","critical level","top talent flag"]},
    "Critical Position":         {type:"string", aliases:["critical position","critical","is critical","critical role","critical flag"]},
    "Job Complexity":            {type:"string", aliases:["job complexity","complexity","job family","job_family"]},

    "Business Unit ID":          {type:"string", aliases:["business unit id","bu id","buid"]},
    "Business Unit Name":        {type:"string", aliases:["business unit name","business unit","bu","bu name","business","business_unit","business area","business pillar"]},
    "Directorate ID":            {type:"string", aliases:["directorate id","dir id"]},
    "Directorate Name":          {type:"string", aliases:["directorate name","directorate","division directorate"]},
    "Departement ID":            {type:"string", aliases:["departement id","department id","dept id"]},
    "Departement Name":          {type:"string", aliases:["departement name","department name","department","dept","dept name","sub department"]},
    "Division ID":               {type:"string", aliases:["division id","div id"]},
    "Division":                  {type:"string", aliases:["division","div","divisi"]},
    "Group ID":                  {type:"string", aliases:["group id","grp id"]},
    "Group Name":                {type:"string", aliases:["group name","group","grp name"]},
    "Pillar Name":               {type:"string", aliases:["pillar name","pillar","business pillar"]},
    "Working Location":          {type:"string", aliases:["working location","work location","location","venue","office","site"]},

    "Supervisor ID":             {type:"string", aliases:["supervisor id","manager id","spv id","line manager id","reports to id","manager employee code"]},
    "Supervisor Name":           {type:"string", aliases:["supervisor name","manager name","spv name","line manager","reports to"]},
    "SPV Position Name":         {type:"string", aliases:["spv position name","supervisor position name","manager position","manager title"]},

    // ── PERSONAL / DEMOGRAPHICS ──
    "Group Age":                 {type:"string", aliases:["group age","age group","age band","age range","age category"]},
    "Years of Experience":       {type:"numeric", aliases:["years of experience","yoe","experience","years experience","tenure","service years"]},
    "Start Date":                {type:"date", aliases:["start date","join date","joining date","hire date","date of joining","doj","start work date"]},
    "End Date":                  {type:"date", aliases:["end date","leave date","resign date","termination date"]},
    "Last Promotion Date":       {type:"date", aliases:["last promotion date","promotion date","last promoted"]},
    "Last Rotation Date":        {type:"date", aliases:["last rotation date","rotation date","last rotated"]},

    // ── TALENT / READINESS ──
    "Final Score Readiness":      {type:"numeric", aliases:["final score readiness","readiness score","score readiness"]},
    "Final Score Readiness Text": {type:"string", validator:"readiness", aliases:["final score readiness text","readiness","readiness text","readiness level","readiness status","successor readiness"]},
    "As Talent Calibrated":       {type:"string", aliases:["as talent calibrated","talent calibrated","calibrated as talent","talent status","is talent"]},
    "Calibration Status":         {type:"string", aliases:["calibration status","calibration","calib status"]},
    "Successor As Talent":        {type:"string", aliases:["successor as talent","successor talent status"]},
    "Incumbent Flight Risk":      {type:"string", validator:"flightrisk", aliases:["incumbent flight risk","flight risk","incumbent risk","attrition risk"]},
    "Successor Flight Risk":      {type:"string", validator:"flightrisk", aliases:["successor flight risk","successor risk"]},

    // ── SUCCESSION ──
    "Successor Name":            {type:"string", aliases:["successor name","succ name","successor","candidate name"]},
    "Successor Position":        {type:"string", aliases:["successor position","successor position name","succ position"]},
    "Successor Position ID":     {type:"string", aliases:["successor position id","succ position id"]},
    "Succession Emp Number":     {type:"string", aliases:["succession emp number","successor emp number","successor nik","successor employee id","succ nik"]},
    "Job Family Successor":      {type:"string", aliases:["job family successor","job family","family"]},
    "Level Successor":           {type:"string", aliases:["level successor","successor level"]},
    "Business Unit Successor":   {type:"string", aliases:["business unit successor","successor business unit","successor bu"]},

    // ── SKILLS / DEVELOPMENT (MyLearning + manual PCD) ──
    "Technical Knowledge":       {type:"string", aliases:["technical knowledge","tech knowledge","technical skill","tech skill"]},
    "Softskill Knowledge":       {type:"string", aliases:["softskill knowledge","soft skill knowledge","soft skills","softskills"]},
    "Certification Plan":        {type:"string", aliases:["certification plan","certifications","certification","certs","certificate plan","license_certif_no","license certif no","license certificate no","license/certification no","certificate number","cert no"]},
    "Training Plan":             {type:"string", aliases:["training plan","trainings","training","program name","training_by_company"]},
    "Training Name":             {type:"string", aliases:["training name","course name","training course","module name","module","nama_training","nama training","program name","modulename"]},
    "Training Provider":         {type:"string", aliases:["training provider","provider","nama_provider","nama provider","module author name","institution","university partner","institution / university partner","institution/university partner","trainer"]},
    "Training Type":             {type:"string", aliases:["training type","jenis_training","jenis training","jenis_training_t","module type","training method","method","subtype name","subtype_name"]},
    "Training Category":         {type:"string", aliases:["training category","kategori","kategori_training","kategori training","category","program category","program category name","primary tag"]},
    "Training Status":           {type:"string", aliases:["training status","module status","completion status","is complete","status_training","training status"]},
    "Completion Percentage":     {type:"percentage", aliases:["completion percentage","completion %","completion pct","percent complete","percent completion","progress","progress %","completion_pct"]},
    "Training Date":             {type:"date", aliases:["training date","tanggal_training","tanggal training","completed on","completion date","date completed","start_date_training","start date training","course period start date","started on","enrolled on"]},
    "Training Hours":            {type:"numeric", aliases:["training hours","learning_hours","learning hours","duration hours","jam pelatihan","time spent (mins)","time spent","time_spent","estimated duration (mins)","estimated duration","duration (mins)"]},
    "Skill Name":                {type:"string", aliases:["skill name","skill","skill name / skills","skills","skill_name","skills to be developed","competency","competency/ skills to be developed"]},
    "Skill Type":                {type:"string", aliases:["skill type","competency type","skill category","skill_type"]},
    "Project Assignment":        {type:"string", aliases:["project assignment","projects","assignments","project assignments"]},
    "Mentor":                    {type:"string", aliases:["mentor","mentor name","assigned mentor"]},
    "Coach":                     {type:"string", aliases:["coach","coach name","assigned coach"]},
    "Mentor Nomination 1 Name":  {type:"string", aliases:["mentor nomination 1 name","mentor nom 1","mentor 1"]},
    "Mentor Nomination 2 Name":  {type:"string", aliases:["mentor nomination 2 name","mentor nom 2","mentor 2"]},
    "Char255":                   {type:"string", aliases:["char255","note","notes","remarks","comment","comments"]}
  }
};

// ── Persistence: hydrate admin overrides + expose save/load helpers ───────
window.MAPPING_CONFIG.loadOverrides = function(sourceId) {
  try {
    const raw = localStorage.getItem(window.MAPPING_CONFIG.storageKey);
    if (!raw) return null;
    const all = JSON.parse(raw);
    return all[sourceId] || null;
  } catch(e){ return null; }
};
window.MAPPING_CONFIG.saveOverrides = function(sourceId, mapping) {
  try {
    const raw = localStorage.getItem(window.MAPPING_CONFIG.storageKey);
    const all = raw ? JSON.parse(raw) : {};
    all[sourceId] = mapping;
    localStorage.setItem(window.MAPPING_CONFIG.storageKey, JSON.stringify(all));
    return true;
  } catch(e){ return false; }
};
window.MAPPING_CONFIG.clearOverrides = function(sourceId) {
  try {
    const raw = localStorage.getItem(window.MAPPING_CONFIG.storageKey);
    if (!raw) return;
    const all = JSON.parse(raw);
    if (sourceId) delete all[sourceId];
    else Object.keys(all).forEach(k => delete all[k]);
    localStorage.setItem(window.MAPPING_CONFIG.storageKey, JSON.stringify(all));
  } catch(e){}
};


const TEMPLATE_B64 = "UEsDBBQAAAAIAIAinVxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIAIAinVyWv9d47gAAACsCAAARAAAAZG9jUHJvcHMvY29yZS54bWzNksFqwzAMhl9l+J7IcUthJvVlY6cNBits7GZktTWLE2NrJH37JVmbMrYH2NHS70+fQDVGjV2i59RFSuwp3wyhabPGuBVH5qgBMh4p2FyOiXZs7rsULI/PdIBo8cMeCJSUGwjE1lm2MAGLuBCFqR1qTGS5S2e8wwUfP1MzwxwCNRSo5QxVWYEw08R4GpoaroAJxpRC/i6QW4hz9U/s3AFxTg7ZL6m+78t+NefGHSp4e3p8mdctfJvZtkjjr+w1nyJtxWXy6+rufvcgjJJqU8h1oW53cq2V1LJ6n1x/+F2FQ+f83v9j44ugqeHXXZgvUEsDBBQAAAAIAIAinVyZXJwjEAYAAJwnAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbO1aW3PaOBR+76/QeGf2bQvGNoG2tBNzaXbbtJmE7U4fhRFYjWx5ZJGEf79HNhDLlg3tkk26mzwELOn7zkVH5+g4efPuLmLohoiU8nhg2S/b1ru3L97gVzIkEUEwGaev8MAKpUxetVppAMM4fckTEsPcgosIS3gUy9Zc4FsaLyPW6rTb3VaEaWyhGEdkYH1eLGhA0FRRWm9fILTlHzP4FctUjWWjARNXQSa5iLTy+WzF/NrePmXP6TodMoFuMBtYIH/Ob6fkTlqI4VTCxMBqZz9Wa8fR0kiAgsl9lAW6Sfaj0xUIMg07Op1YznZ89sTtn4zK2nQ0bRrg4/F4OLbL0otwHATgUbuewp30bL+kQQm0o2nQZNj22q6RpqqNU0/T933f65tonAqNW0/Ta3fd046Jxq3QeA2+8U+Hw66JxqvQdOtpJif9rmuk6RZoQkbj63oSFbXlQNMgAFhwdtbM0gOWXin6dZQa2R273UFc8FjuOYkR/sbFBNZp0hmWNEZynZAFDgA3xNFMUHyvQbaK4MKS0lyQ1s8ptVAaCJrIgfVHgiHF3K/99Ze7yaQzep19Os5rlH9pqwGn7bubz5P8c+jkn6eT101CznC8LAnx+yNbYYcnbjsTcjocZ0J8z/b2kaUlMs/v+QrrTjxnH1aWsF3Pz+SejHIju932WH32T0duI9epwLMi15RGJEWfyC265BE4tUkNMhM/CJ2GmGpQHAKkCTGWoYb4tMasEeATfbe+CMjfjYj3q2+aPVehWEnahPgQRhrinHPmc9Fs+welRtH2Vbzco5dYFQGXGN80qjUsxdZ4lcDxrZw8HRMSzZQLBkGGlyQmEqk5fk1IE/4rpdr+nNNA8JQvJPpKkY9psyOndCbN6DMawUavG3WHaNI8ev4F+Zw1ChyRGx0CZxuzRiGEabvwHq8kjpqtwhErQj5iGTYacrUWgbZxqYRgWhLG0XhO0rQR/FmsNZM+YMjszZF1ztaRDhGSXjdCPmLOi5ARvx6GOEqa7aJxWAT9nl7DScHogstm/bh+htUzbCyO90fUF0rkDyanP+kyNAejmlkJvYRWap+qhzQ+qB4yCgXxuR4+5Xp4CjeWxrxQroJ7Af/R2jfCq/iCwDl/Ln3Ppe+59D2h0rc3I31nwdOLW95GblvE+64x2tc0LihjV3LNyMdUr5Mp2DmfwOz9aD6e8e362SSEr5pZLSMWkEuBs0EkuPyLyvAqxAnoZFslCctU02U3ihKeQhtu6VP1SpXX5a+5KLg8W+Tpr6F0PizP+Txf57TNCzNDt3JL6raUvrUmOEr0scxwTh7LDDtnPJIdtnegHTX79l125COlMFOXQ7gaQr4Dbbqd3Do4npiRuQrTUpBvw/npxXga4jnZBLl9mFdt59jR0fvnwVGwo+88lh3HiPKiIe6hhpjPw0OHeXtfmGeVxlA0FG1srCQsRrdguNfxLBTgZGAtoAeDr1EC8lJVYDFbxgMrkKJ8TIxF6HDnl1xf49GS49umZbVuryl3GW0iUjnCaZgTZ6vK3mWxwVUdz1Vb8rC+aj20FU7P/lmtyJ8MEU4WCxJIY5QXpkqi8xlTvucrScRVOL9FM7YSlxi84+bHcU5TuBJ2tg8CMrm7Oal6ZTFnpvLfLQwJLFuIWRLiTV3t1eebnK56Inb6l3fBYPL9cMlHD+U751/0XUOufvbd4/pukztITJx5xREBdEUCI5UcBhYXMuRQ7pKQBhMBzZTJRPACgmSmHICY+gu98gy5KRXOrT45f0Usg4ZOXtIlEhSKsAwFIRdy4+/vk2p3jNf6LIFthFQyZNUXykOJwT0zckPYVCXzrtomC4Xb4lTNuxq+JmBLw3punS0n/9te1D20Fz1G86OZ4B6zh3OberjCRaz/WNYe+TLfOXDbOt4DXuYTLEOkfsF9ioqAEativrqvT/klnDu0e/GBIJv81tuk9t3gDHzUq1qlZCsRP0sHfB+SBmOMW/Q0X48UYq2msa3G2jEMeYBY8wyhZjjfh0WaGjPVi6w5jQpvQdVA5T/b1A1o9g00HJEFXjGZtjaj5E4KPNz+7w2wwsSO4e2LvwFQSwMEFAAAAAgAgCKdXPqCk27VCwAAFV8AABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWydnN122lgShV9Fi/tpUwIh0cv2WjY/7nQnGY+dpKfnToZjWxOBaEnEyTz9SIJzwE7VqW3fdAeb2pRqA7v0IXz6VJRfq0dj6uD7Kl9XZ73Hut78enJSLR7NKq1+KTZm3fzmvihXad3cLB9Oqk1p0mVXtMpPwn5/dLJKs3Xv/LT72XV5flps6zxbm+syqLarVVr+uDR58XTWo579wU328Fi3Pzg5P92kD+bW1J8312Vz68SpLLOVWVdZsQ5Kc3/Wu6Bfr6ZdQXePL5l5qo7+HbSHclcUX9sb75ZnvX6vlV6b4MftJs+6BwvqYvPe3NcTk+eNYNgL0kWdfTPXzd3OendFXRer9vdNm3VaNz+6L4v/mXX3mCY3zX2bZjY/3Xknshdtj/HvfcM9dzxtU8f/tp3Pu8E2g7pLKzMp8j+zZf141kt6wdLcp9u8vimefjP7YUWt3qLIq+6/wdPuvmFzGItt1XSzL246WGXr3f/T7/shIwXhviBECwb7ggFaMNwXDNGCaF8QoQWjfcEILYj3BTFakOwLErRgvC8YowXUt8714RJnNuw2WbsJ9pus4QQ7TtZygj0nazrBrpO1nWDfyRpPsPNkrSfYe7LmE+x+aN0PYfdD636Iv9bdix12P7Tuh7D7oXU/hN0Prfsh7H5o3Q9h90Prfgi7H1r3Q9j90Lofwu4PrPsD2P2BdX8Auz+w7g/w93r3Zg+7P7DuD2D3B9b9Aez+wLo/gN0fWPcHsPsD6/4Adn9g3R/A7g+t+0PY/aF1fwi7P7TuD2H3h9b9IZ71Luxh94fW/SHs/tC6P4TdH1r3h7D7Q+v+EHZ/aN0fwu5H1v0Idj+y7kew+5F1P4Ldj6z7Eex+ZN2P8F3PLXuw+5F1P4Ldj6z7Eex+ZN2PYPcj634Euz+y7o9g90fW/RHs/si6P4LdH1n3R7D7I+v+CHZ/ZN0f4bu+W/Zh90fW/RHs/si6P4LdH1n3R7D7sXU/ht2Prfsx7H5s3Y9h92Prfgy7H1v3Y9j92Lofw+7H1v0YP9dzJ3uw+7F1P4bdj637Mex+Yt1PYPcT634Cu59Y9xPY/cS6n8DuJ9b9BHY/se4nsPuJdT+B3U+s+wl+ru9O9mH3E+t+Ars/tu6PYffH1v0x7P7Yuj+G3R9b98ew+2Pr/hh2f2zdH8Puj637Y9j9sXV/DLs/tu6PcdbjYM8raM8B9+C8p++ATx8nPn2HfPo48+k76NPHqU/fYZ8+zn36Dvz0cfLTd+inj7OfvoM/fZz+9B3+6eP8p+8AUB9/RhwA4GsI4AEB4s+IAwR8BQU8YMBXcMADCHwFCTygwFewwAMMfAUNPODAV/DAAxB8BRE8IEGcCZKDgoRTQXJYkHAuSA4MEk4GyaFBwtkgOThIOB0khwcJ54PkACHhhJAcIiScEZKDhIRTQnKYkHBOSA4UEk4KyaFCwlkhOVhIOC2kweGjAvwZ4YAh4cSQHDIknBmSg4aEU0Ny2JBwbkgOHBJODsmhQ8LZITl4SDg9JIcPCeeH5AAi4QSRHEIknCHS8PDxEf6McBiRcI5IDiQSThLJoUTCWSI5mEg4TSSHEwnnieSAIuFEkRxSJJwpkoOKhFNFcliRcK5IDiwSThYpOnykiD8jHFwknC6Sw4uE80VygJFwwkgOMRLOGMlBRsIpIznMSDhnJAcaCSeN5FAj4ayRHGwknDaSw42E80YaHT5mxp8RDjkSzhzJQUfCqSM57Eg4dyQHHgknj+TQI+HskRx8JJw+ksOPhPNHcgCScAJJDkESziDJQUjCKSTFh0sP8GeEA5GEk0hyKJJwFkkORhJOI8nhSMJ5JDkgSTiRJIckCWeS5KAk4VSSHJYknEuSA5O0I5Mni+6Cre5qr2lap+enZfEUlF1Fe1VXu8fvlNx1Xr3modp7dNeSdXdsfpqt20vqbuuy+W3WCNbnH9/9cXpSN4/Q3jxZ7Isu/UWz1SYvfhgTfExXhimf+MuvymK7CS4euNKpv/S2Tss6aEbA1c6UrtdLqXLur7wuqqy7bO/dlCm+AouFYf0GlndTY+rf+et/L+6CebrK8h9M7e/+2sk/3ptvJmcK/1AKy6bnRZoHtntG4j143I3l9bZiBD4oAlmepyVv2UeoVDDsn/7i92lVB9dlsSq69oVn3DUgclM0h+7R+BfyMmOP/wapFA7/1l97ua2an1RV8Hmd1fyjf3qNgtDFZ7/GNCvNoi7KZnB8D1/weqGDPxUFs2neqszKrIUp/BuvFzr4SzuGb1klvmv9BytmKi+USJlsy7LtWnrvuFDS5Xa7MWXz6IXw4r1Q4uWoXhjchZYy118C/3v+hZI1zxSkLpTU+WQWj+vuXfT2a/OGFNyYv7fN03IZECempJAsFkxNteAUlWB6qUjBp9JUj0W+5MSUlJLFxPaU7BIPOOTElDyTxcT2lHh7qRh6p6dEnSwmtqckoHjAA05MSURZTGxPiceXigPv9JSclMXE9pT8FA94yIkpgSqLie0pAftSceidnpK0spjYnpK94gFHnJgSw7KY2J4SzC8VI+/0lIyWxcT2lOC+Le5rJDAulRRndaSmLrVIZ8U8c7vUMt4rKLap5T6jysXFpZb+nI7YlLIG8GK+2SmrwOQxLcOIezFdKpHPtcIFwqWS9qyOOB8l6Hkx33yUsPcLim0qgc+pcnFwqWQ9qyM2pcQ8L+abnRL1fkGxTSXuOVX2+askPasjNqWEPC/mm50S9H5BsU0l7DnVEaej5DyrIzalRDwv5pudEvN+QbFNJeo51ZjTUVKe1RGbekPAx77ZTd6Q9LE6u8kbIj/hdN6Q9InY1BsCPvHO7g1Jn+izUyL/3XqxXd21DGWetzg/uMmqr5yOkvS328XCVB0Dmq02wcdWtOR0lNy/qIJPad72M0nz7K6lYuywtNjftSNzmYkW9U7AT2cmWsL/pMOJKIF+AOuB0+NklDzvKJlfQQnx57zUq6Slt5uLc5xT0dLaqbRnNxbocEJaXLNC4otKy2te7dOj59WvJbYiKbaqZTany52ITLTQZoXEtrTU5tW8E9RyW5EUW9WSm9PlTlUmWnSzQmJbWnbzar4JTrX0ViSlVqcqhWd0uROWqYrjOSGxLS3BeTXvBLUMVyTFVrUTd06XO22ZYjH+QkhsSzuJ59W8E4SDnZcUW4Xz/mjP4pJkCgf+sZDYlnYqz6t5J6id1SuSYqvaCT6nyyXJFF4NjoXEtuAV4Zmad4LwssBLiq3Ca8ORLpckU3hZOBYS24JXhGdq3gnCywIvKbYKrw1HumySwMvCsZDYFrwiPFPzThBeFnhJsVV4bTjS5ZJkBi8Lx0JSWzN4RXim5pvgDF4WeEmxVXhtONLlMNgMXhaOhcS24BXhmZp3gvCywEuKrcJrw5EuB8Nm8LJwLCS29ZYVIfZP8C3LQgxM8C1rA4fEZm9ZFkQmNnvLipD4J/iWZSEBJgivDRfLZcdt0lyDDDN4bxA0xWbhxaERxpnDDFgjPIpit/A+wY+B2xln8EIhaIrNwhvFi0H4V8gZsF94FMVu4UWDHwO3TM7gTUPQFJuFV40Xg/DvlnNg8fAoSt3O4Q2EHwO3Zc7hFUTQFJuFd5AXg/AvnXNgI/Eoit3Cqwk/Bm79nMO7iaApNgsvJy8G4d9G58Cq4lEUu9Uuqy/SxSNXpywmH8y6Zrn+XLuWviz+axZ1cFFV2cN6xWP9ubJ1TExZZ/fZYndJ+nWech+azLXrCss0W2frB+kToLl2KaGtlx5fu3BQe3ztWkHt8bUrA7XHV9J/9xQIPharbL0zgviPwOZK3nNCUk9KzP8sFQo9KbHOCUk9KVHefjg3KVab3HzPau6bL3MltP8yaVkFxX0w+74xZWbWC7YN+PK+P9bFU26W7Hee5sAn/lX35uNTuVIieJ6177a3i6I0wY1Jl91nhpyOkrmsTvMO+Z17U7mCw9b/SfeVkq32c2npC0MnZeH+dPruC3TtH4z/kJYP2boKcnPfqPZ/ab+9We6+P7e7UReb7jF3f6i9++djc8imbO/Q/P6+KGp7o/2anvtL+Of/B1BLAwQUAAAACACAIp1ccnZobK8CAABFCwAADQAAAHhsL3N0eWxlcy54bWzdVtuK2zAQ/RXjD6g3MTVxSQKtYaHQloXdh74qsZwIdHFleevs13dGcuxko1m2faxDkGaOZs7cFGfduZPkj0fOXTIoqbtNenSu/ZRl3f7IFes+mJZrQBpjFXMg2kPWtZazukMjJbPl3V2RKSZ0ul3rXt0r1yV702u3Se/SbLtujJ41yzQo4ChTPHlmcpNWTIqdFf4sU0KegnqJir2RxiYOQuGbdIGa7iXAiyBhlKMfJbSxqMwCwxs8u9HHTGEPO4h3fG55Rpd+6cC1kPI6J1Bs1y1zjlt9D4K38cobKBn3T6cWkjpYdlosP6azgV+AZGdsze0VTVBt15I3DgysOBxxdabFpIxzRsGmFuxgNPMxnC3mFQ8lvu1QMl6LXqVkEbJw/MLPuIHw9lzKR3Tzs5liXICroUnCGHytcQISLNl5C4mN2+AmCEh06S34vnC7/Ce3SSuejfvSQwbay7964/iD5Y0YvDw0Ez/lfUF4Bz1rW3n6LMVBKx5yfzfhds3OdsnRWPECbDhre1BwmybP3Dqxv9T8tqx94oMbZzYbmrFmU7l88a4aMWkTvAOb9AdeYTlHkex6IZ3Qo3QUdc31TT/AvWM7+I248g/na96wXrqnCdyk8/67H6xyOvWAlRlPzftvOL+LYrqFwCV0zQdeV6MIE1nFR/M1cu+fOELZBCyOIEbxUBFQNsGK4vmf8lmR+QSMim0VRVakzYq0CVYxpPIfiiduU8ITz7Qs87woqIpWVTSCiqpbUeA37o2KDS0oHmT6u1rT3aYn5O05oHr61oRQmdKTSGVK1xqReN3Qoizj3aZ40ILqAjU7yB/nwZmK2+Q5dpWKjbrBNFKWFIKzGJ/RoiCqU+An3h/qluR5WcYRxOIR5DmF4G2kESoCjIFC8ty/B1+9j7Lzeyqb/zhv/wBQSwMEFAAAAAgAgCKdXJeKuxzAAAAAEwIAAAsAAABfcmVscy8ucmVsc52SuW7DMAxAf8XQnjAH0CGIM2XxFgT5AVaiD9gSBYpFnb+v2qVxkAsZeT08EtweaUDtOKS2i6kY/RBSaVrVuAFItiWPac6RQq7ULB41h9JARNtjQ7BaLD5ALhlmt71kFqdzpFeIXNedpT3bL09Bb4CvOkxxQmlISzMO8M3SfzL38ww1ReVKI5VbGnjT5f524EnRoSJYFppFydOiHaV/Hcf2kNPpr2MitHpb6PlxaFQKjtxjJYxxYrT+NYLJD+x+AFBLAwQUAAAACACAIp1cGrobqzABAAAjAgAADwAAAHhsL3dvcmtib29rLnhtbI1R0UrDQBD8lXAfYFLRgqXpi0UtiBYrfb8km2bp3W3Y27Tar3eTECz44tPezizDzNzyTHwsiI7Jl3ch5qYRaRdpGssGvI031EJQpib2VnTlQxpbBlvFBkC8S2+zbJ56i8GslpPWltPrhQRKQQoK9sAe4Rx/+X5NThixQIfynZvh7cAkHgN6vECVm8wksaHzCzFeKIh1u5LJudzMRmIPLFj+gXe9yU9bxAERW3xYNZKbeaaCNXKU4WLQt+rxBHo8bp3QEzoBXluBZ6auxXDoZTRFehVj6GGaY4kL/k+NVNdYwprKzkOQsUcG1xsMscE2miRYD7kZLA6BdG6qMZyoq6uqeIFK8KYa/U2mKqgxQPWmOlFxLajcctKPQef27n72oEV0zj0q9h5eyVZTxul/Vj9QSwMEFAAAAAgAgCKdXCQem6KtAAAA+AEAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc7WRPQ6DMAyFrxLlADVQqUMFTF1YKy4QBfMjEhLFrgq3L4UBkDp0YbKeLX/vyU6faBR3bqC28yRGawbKZMvs7wCkW7SKLs7jME9qF6ziWYYGvNK9ahCSKLpB2DNknu6Zopw8/kN0dd1pfDj9sjjwDzC8XeipRWQpShUa5EzCaLY2wVLiy0yWoqgyGYoqlnBaIOLJIG1pVn2wT06053kXN/dFrs3jCa7fDHB4dP4BUEsDBBQAAAAIAIAinVxlkHmSGQEAAM8DAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK2TTU7DMBCFrxJlWyUuLFigphtgC11wAWNPGqv+k2da0tszTtpKoBIVhU2seN68z56XrN6PEbDonfXYlB1RfBQCVQdOYh0ieK60ITlJ/Jq2Ikq1k1sQ98vlg1DBE3iqKHuU69UztHJvqXjpeRtN8E2ZwGJZPI3CzGpKGaM1ShLXxcHrH5TqRKi5c9BgZyIuWFCKq4Rc+R1w6ns7QEpGQ7GRiV6lY5XorUA6WsB62uLKGUPbGgU6qL3jlhpjAqmxAyBn69F0MU0mnjCMz7vZ/MFmCsjKTQoRObEEf8edI8ndVWQjSGSmr3ghsvXs+0FOW4O+kc3j/QxpN+SBYljmz/h7xhf/G87xEcLuvz+xvNZOGn/mi+E/Xn8BUEsBAhQDFAAAAAgAgCKdXEbHTUiVAAAAzQAAABAAAAAAAAAAAAAAAIABAAAAAGRvY1Byb3BzL2FwcC54bWxQSwECFAMUAAAACACAIp1clr/XeO4AAAArAgAAEQAAAAAAAAAAAAAAgAHDAAAAZG9jUHJvcHMvY29yZS54bWxQSwECFAMUAAAACACAIp1cmVycIxAGAACcJwAAEwAAAAAAAAAAAAAAgAHgAQAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQIUAxQAAAAIAIAinVz6gpNu1QsAABVfAAAYAAAAAAAAAAAAAACAgSEIAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECFAMUAAAACACAIp1ccnZobK8CAABFCwAADQAAAAAAAAAAAAAAgAEsFAAAeGwvc3R5bGVzLnhtbFBLAQIUAxQAAAAIAIAinVyXirscwAAAABMCAAALAAAAAAAAAAAAAACAAQYXAABfcmVscy8ucmVsc1BLAQIUAxQAAAAIAIAinVwauhurMAEAACMCAAAPAAAAAAAAAAAAAACAAe8XAAB4bC93b3JrYm9vay54bWxQSwECFAMUAAAACACAIp1cJB6boq0AAAD4AQAAGgAAAAAAAAAAAAAAgAFMGQAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAMUAAAACACAIp1cZZB5khkBAADPAwAAEwAAAAAAAAAAAAAAgAExGgAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLBQYAAAAACQAJAD4CAAB7GwAAAAA=";

function downloadTemplate(){
  const bytes=Uint8Array.from(atob(TEMPLATE_B64),c=>c.charCodeAt(0));
  const blob=new Blob([bytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);
  a.download="succession_template_data.xlsx";a.click();
}

// Prevent datalabels from auto-activating on ALL charts (only use where plugins:[ChartDataLabels] is set)
if(typeof ChartDataLabels!=="undefined")Chart.unregister(ChartDataLabels);

// ═══════════════════════════════════════════════════════════════════════════
// MODULAR MULTI-SOURCE UPLOAD ENGINE
// ═══════════════════════════════════════════════════════════════════════════
//
// Architecture layers:
//   1. FIELD_REGISTRY    : canonical fields + aliases + validators (single source of truth)
//   2. SOURCE_REGISTRY   : per-system metadata (required fields, hints)
//   3. Column Detector   : header → canonical field auto-mapping
//   4. Manual Mapper UI  : admin maps unrecognized columns
//   5. Validation Engine : per-source rules + cross-source rules
//   6. Integration Layer : merges all sources → ALL[] for the dashboard
//
// Extending: add a new HR system by appending an entry to SOURCE_REGISTRY,
// and add any new canonical fields to FIELD_REGISTRY.

// ── Canonical Field Registry ───────────────────────────────────────────────
// Each entry: { canonical name → { aliases:[...], validator:fn, required?, critical? } }
// `canonical` is the key written into the merged row that the dashboard reads.
// ── FIELD_REGISTRY ─────────────────────────────────────────────────────────
// Loaded from window.MAPPING_CONFIG (separate <script> block at top of file).
// This indirection keeps mapping configuration DECOUPLED from dashboard logic:
// admins can edit MAPPING_CONFIG (or override via the UI / localStorage) without
// touching the engine.
const FIELD_REGISTRY = window.MAPPING_CONFIG.fields;

// Build a fast lookup: any lowercased-normalized alias → canonical field
const _aliasIndex = (() => {
  const idx = {};
  for (const canon in FIELD_REGISTRY) {
    const entry = FIELD_REGISTRY[canon];
    const all = [canon, ...(entry.aliases || [])];
    all.forEach(a => { idx[_norm(a)] = canon; });
  }
  return idx;
})();

function _norm(s){return String(s||"").toLowerCase().replace(/[\s_\-./\\:]+/g," ").replace(/\s+/g," ").trim();}
function _fuzzyMatch(header){
  // exact alias hit
  const exact = _aliasIndex[_norm(header)];
  if (exact) return {canonical: exact, confidence: "auto"};
  // substring containment: canonical name contained in header, or vice versa
  const h = _norm(header);
  let best = null;
  for (const alias in _aliasIndex) {
    if (alias.length < 3) continue;
    if (h.includes(alias) || alias.includes(h)) {
      const score = Math.min(alias.length, h.length) / Math.max(alias.length, h.length);
      if (!best || score > best.score) best = {canonical: _aliasIndex[alias], score};
    }
  }
  if (best && best.score > 0.6) return {canonical: best.canonical, confidence: "fuzzy"};
  return null;
}

// ── Source Registry ───────────────────────────────────────────────────────
// Each source: id, label, description, required canonical fields, icon.
// To add a new HR system, append an entry here.
const SOURCE_REGISTRY = [
  {
    id: "sap",
    label: "SAP Employee Master",
    desc: "Core employee data: NIK, position, department, org structure",
    required: ["NIK","Employee Name"],
    primary: true,
    icon: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 12h8M8 8h8M8 16h5"/></svg>'
  },
  {
    id: "mylearning",
    label: "MyLearning",
    desc: "Certifications, training records, learning history",
    required: [],
    primary: false,
    icon: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>'
  }
];

// ── Source State ──────────────────────────────────────────────────────────
// Per-source: file, raw rows, header→canonical mapping, validation issues, status
const SOURCES = {};
SOURCE_REGISTRY.forEach(s => SOURCES[s.id] = {
  id: s.id, file: null, raw: [], headers: [], mapping: {}, issues: [], status: "empty"
});

let ALL = [], FILTERED = [], CHARTS = {}, PF_STATE = "all";
let _pendingSource = null;  // { sourceId, file, raw, headers, autoMapping }

const VR=["ready now : less than a year","ready later : 1-3 years","ready future : 3-5 years",""];
const VF=["low","medium","high","unknown",""];

// ── Cell-level validators (referenced by FIELD_REGISTRY) ──────────────────
const VALIDATORS = {
  readiness: v => { const s = String(v||"").toLowerCase().trim(); return !s || VR.includes(s); },
  flightrisk: v => { const s = String(v||"").toLowerCase().trim(); return !s || VF.includes(s); }
};

// ── Type validators (referenced by FIELD_REGISTRY.fields[X].type) ────────
// Empty / dash / nan are always tolerated (treated as "not provided").
const TYPE_VALIDATORS = {
  string:     v => true,
  numeric:    v => { const s = String(v||"").trim(); if(!s||s==="-"||s.toLowerCase()==="nan") return true; return !isNaN(Number(s.replace(/,/g,""))); },
  percentage: v => { const s = String(v||"").trim().replace(/%$/,""); if(!s||s==="-"||s.toLowerCase()==="nan") return true; const n=Number(s); return !isNaN(n) && n>=0 && n<=100; },
  date:       v => {
    const s = String(v||"").trim(); if(!s||s==="-"||s.toLowerCase()==="nan") return true;
    // Accept ISO, dd/mm/yyyy, dd-mm-yyyy, "DD Mon YYYY", and Excel serial numbers
    if (/^\d{4,5}(\.\d+)?$/.test(s)) return true;  // Excel serial date
    if (/^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$/.test(s)) return true;
    if (/^\d{4}-\d{1,2}-\d{1,2}/.test(s)) return true;
    if (/^\d{1,2}\s+[A-Za-z]{3,}\s+\d{2,4}$/.test(s)) return true;
    return !isNaN(new Date(s).getTime());
  },
  boolean:    v => { const s = String(v||"").toLowerCase().trim(); return !s || ["yes","no","y","n","true","false","1","0"].includes(s); }
};

// ═══════════════════════════════════════════════════════════════════════════
// UPLOAD HUB UI
// ═══════════════════════════════════════════════════════════════════════════
function renderHub(){
  const wrap = document.getElementById("hubSources");
  wrap.innerHTML = "";
  SOURCE_REGISTRY.forEach(srcMeta => {
    const s = SOURCES[srcMeta.id];
    const loaded = s.status === "ready" || s.status === "warn";
    const errCount = s.issues.filter(i => i.severity === "CRITICAL" || i.severity === "ERROR").length;
    const warnCount = s.issues.filter(i => i.severity === "WARN").length;

    const card = document.createElement("div");
    card.className = "src-card" + (loaded ? " loaded" : "") + (srcMeta.primary ? " required" : "");
    card.dataset.sid = srcMeta.id;
    card.innerHTML = `
      <div class="src-head">
        <div class="src-icn">${srcMeta.icon}</div>
        <div class="src-info">
          <div class="src-name">${srcMeta.label}</div>
          <div class="src-desc">${srcMeta.desc}</div>
        </div>
      </div>
      <div class="src-stats">
        ${loaded ? `<span class="ok">✓ ${s.raw.length} rows</span>` : `<span>No file</span>`}
        ${loaded ? `<span class="ok">${Object.keys(s.mapping).filter(h=>s.mapping[h]).length} identified columns</span>` : ""}
        ${loaded ? `<span>${(s.headers||[]).length - Object.keys(s.mapping).filter(h=>s.mapping[h]).length} unknown columns</span>` : ""}
      </div>
      ${s.file ? `<div class="src-fname" title="${s.file.name}">${s.file.name}</div>` : `<div class="src-fname" style="visibility:hidden">.</div>`}
      <div class="src-actions">
        <button class="src-btn" onclick="${srcMeta.id==='mylearning' ? `openMyLearningDualUpload()` : `pickFileFor('${srcMeta.id}')`}">
          <svg viewBox="0 0 24 24" fill="none"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
          ${loaded ? "Replace" : "Upload"}
        </button>
        ${loaded ? `<button class="src-btn" onclick="openMapperFor('${srcMeta.id}')">
          <svg viewBox="0 0 24 24" fill="none"><path d="M3 6h18M3 12h18M3 18h18"/></svg>
          Map
        </button>
        <button class="src-btn danger" onclick="clearSource('${srcMeta.id}')">
          <svg viewBox="0 0 24 24" fill="none"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg>
        </button>` : ""}
      </div>`;
    // Drag-and-drop per card
    card.addEventListener("dragover", e => { e.preventDefault(); card.classList.add("drag"); });
    card.addEventListener("dragleave", () => card.classList.remove("drag"));
    card.addEventListener("drop", e => {
      e.preventDefault(); card.classList.remove("drag");
      if (e.dataTransfer.files[0]) ingestFile(srcMeta.id, e.dataTransfer.files[0]);
    });
    wrap.appendChild(card);
  });
  refreshHubSummary();
}

function refreshHubSummary(){
  const loadedCount = Object.values(SOURCES).filter(s => s.status === "ready" || s.status === "warn").length;
  const totalRows = Object.values(SOURCES).reduce((a,s) => a + s.raw.length, 0);
  const allIssues = Object.values(SOURCES).flatMap(s => s.issues);
  const issueCount = allIssues.length;

  // Quick employee unique-count preview from loaded sources
  let uniqueEmps = 0;
  if (loadedCount > 0) {
    const ids = new Set();
    Object.values(SOURCES).forEach(s => {
      if (!s.raw.length) return;
      s.raw.forEach(row => {
        const nik = _extractCanonValue(row, s.mapping, "NIK");
        const name = _extractCanonValue(row, s.mapping, "Employee Name");
        const key = nik ? "nik:"+_norm(nik) : (name ? "nm:"+_norm(name) : null);
        if (key) ids.add(key);
      });
    });
    uniqueEmps = ids.size;
  }

  document.getElementById("hsSources").textContent = loadedCount;
  document.getElementById("hsRows").textContent = totalRows;
  document.getElementById("hsEmployees").textContent = uniqueEmps;
  document.getElementById("hsIssues").textContent = issueCount;
  document.getElementById("btnIssues").style.display = issueCount > 0 ? "inline-flex" : "none";

  // Build button: needs the primary source loaded with at least an Employee Name or NIK mapping
  const primary = SOURCES["sap"];
  const hasIdentity = primary && (primary.status === "ready" || primary.status === "warn") &&
    Object.values(primary.mapping).some(v => v === "NIK" || v === "Employee Name");
  // Fallback: allow build if ANY source has an identity column (e.g., user uploaded only TM)
  const anyHasIdentity = Object.values(SOURCES).some(s =>
    (s.status === "ready" || s.status === "warn") &&
    Object.values(s.mapping).some(v => v === "NIK" || v === "Employee Name"));
  document.getElementById("btnBuild").disabled = !(hasIdentity || anyHasIdentity);
}

// ═══════════════════════════════════════════════════════════════════════════
// FILE INGESTION
// ═══════════════════════════════════════════════════════════════════════════
function pickFileFor(sourceId){
  const inp = document.getElementById("fileInput");
  inp.onchange = e => {
    if (e.target.files[0]) ingestFile(sourceId, e.target.files[0]);
    inp.value = "";
    inp.onchange = null;
  };
  inp.click();
}

function ingestFile(sourceId, file){
  if (!/\.(xlsx|xls|xlsm|csv|tsv|txt)$/i.test(file.name)) {
    _alert("Invalid File Type", "File must be .xlsx, .xls, .csv, or .tsv",
      [{severity:"ERROR", msg:file.name+" has an unsupported extension."}]);
    return;
  }
  const r = new FileReader();
  r.onload = ev => {
    try {
      // SheetJS auto-detects format from buffer; handles xlsx/xls/csv/tsv (incl. semicolon)
      const wb = XLSX.read(ev.target.result, {type:"array", raw:false});
      // Skip instruction sheets; take first data sheet
      const idx = wb.SheetNames.findIndex(s => s.toLowerCase() !== "instructions" && s.toLowerCase() !== "instruction");
      const ws = wb.Sheets[wb.SheetNames[idx >= 0 ? idx : 0]];
      const raw = XLSX.utils.sheet_to_json(ws, {defval:"", raw:false});
      if (!raw.length) {
        _alert("Empty File", file.name+" contains no data rows.", [{severity:"ERROR", msg:"Sheet has 0 data rows."}]);
        return;
      }
      const headers = Object.keys(raw[0]);
      // ── Auto-detect canonical mapping for each header ────────────────────
      const autoMapping = {};
      // First, try to hydrate from saved admin overrides
      const savedOverrides = window.MAPPING_CONFIG.loadOverrides
        ? window.MAPPING_CONFIG.loadOverrides(sourceId) : null;
      headers.forEach(h => {
        if (savedOverrides && Object.prototype.hasOwnProperty.call(savedOverrides, h)) {
          // Admin previously set this header → trust their override
          const canon = savedOverrides[h];
          autoMapping[h] = {canonical: canon || null, confidence: canon ? "saved" : "ignored"};
          return;
        }
        const match = _fuzzyMatch(h);
        autoMapping[h] = match ? {canonical: match.canonical, confidence: match.confidence} : {canonical: null, confidence: "unmapped"};
      });
      _pendingSource = {sourceId, file, raw, headers, autoMapping};
      openMapper(_pendingSource, /*isNew=*/true);
    } catch (err) {
      _alert("Read Error", "Could not read "+file.name, [{severity:"ERROR", msg:err.message}]);
    }
  };
  r.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════════════════════════════════════
// COLUMN MAPPER (admin UI for unrecognized columns)
// ═══════════════════════════════════════════════════════════════════════════
function openMapperFor(sourceId){
  const s = SOURCES[sourceId];
  if (!s.raw.length) return;
  // Recreate "pending" snapshot from current mapping for re-editing
  const autoMapping = {};
  s.headers.forEach(h => {
    const canon = s.mapping[h] || null;
    autoMapping[h] = canon
      ? {canonical: canon, confidence: "manual"}
      : (_fuzzyMatch(h) ? _fuzzyMatch(h) : {canonical: null, confidence: "unmapped"});
  });
  _pendingSource = {sourceId, file: s.file, raw: s.raw, headers: s.headers, autoMapping};
  openMapper(_pendingSource, /*isNew=*/false);
}

function openMapper(pending, isNew){
  const meta = SOURCE_REGISTRY.find(m => m.id === pending.sourceId);
  document.getElementById("mapTitle").textContent = "Map Columns — " + meta.label;
  document.getElementById("mapSub").textContent =
    `${pending.file.name} · ${pending.raw.length} rows · ${pending.headers.length} columns. Auto-mappings shown — adjust unknowns or leave as "Ignore".`;

  // Build dropdown option list (sorted, with "Ignore" first)
  const canonOptions = Object.keys(FIELD_REGISTRY).sort();

  // Render rows
  const list = document.getElementById("mapList");
  list.innerHTML = "";
  pending.headers.forEach(h => {
    const m = pending.autoMapping[h];
    const row = document.createElement("div");
    row.className = "map-row";
    const badgeClass = m.canonical ? "b-auto" : "b-ignore";
    const badgeText  = m.canonical ? "Identified" : "Unknown";
    row.innerHTML = `
      <div class="mr-src" title="${_esc(h)}">${_esc(h)}</div>
      <div class="mr-arrow">→</div>
      <select data-header="${_esc(h)}" class="${m.canonical ? 'mapped' : 'unmapped'}" onchange="_updateMapBadge(this)">
        <option value="">— Ignore this column —</option>
        ${canonOptions.map(c => `<option value="${_esc(c)}"${c===m.canonical?" selected":""}>${_esc(c)}</option>`).join("")}
      </select>
      <span class="mr-badge ${badgeClass}">${badgeText}</span>`;
    list.appendChild(row);
  });

  // Summary strip
  const identN = Object.values(pending.autoMapping).filter(m => m.canonical).length;
  const unmapN = Object.values(pending.autoMapping).filter(m => !m.canonical).length;
  document.getElementById("mapSummary").innerHTML = `
    <div class="ms-ok"><strong>${identN}</strong> identified columns</div>
    <div class="ms-warn"><strong>${unmapN}</strong> unknown columns (will be ignored)</div>
    <div><strong>${pending.headers.length}</strong> total columns</div>`;

  document.getElementById("mapperModal").classList.add("show");
}

function _updateMapBadge(sel){
  const row = sel.closest(".map-row");
  const badge = row.querySelector(".mr-badge");
  if (sel.value) {
    sel.classList.remove("unmapped"); sel.classList.add("mapped");
    badge.className = "mr-badge b-auto"; badge.textContent = "Identified";
  } else {
    sel.classList.remove("mapped"); sel.classList.add("unmapped");
    badge.className = "mr-badge b-ignore"; badge.textContent = "Unknown";
  }
}

function closeMapper(){
  document.getElementById("mapperModal").classList.remove("show");
  _pendingSource = null;
}

function confirmMapping(){
  if (!_pendingSource) return;
  const finalMapping = {};
  document.querySelectorAll("#mapList select").forEach(sel => {
    const h = sel.dataset.header;
    if (sel.value) finalMapping[h] = sel.value;
  });

  // Commit to SOURCES
  const s = SOURCES[_pendingSource.sourceId];
  s.file = _pendingSource.file;
  s.raw = _pendingSource.raw;
  s.headers = _pendingSource.headers;
  s.mapping = finalMapping;

  // ── Persist admin's mapping overrides to localStorage (per source) ─────
  // Record ALL headers, including those left blank (= explicit "ignore")
  if (window.MAPPING_CONFIG && window.MAPPING_CONFIG.saveOverrides) {
    const persisted = {};
    _pendingSource.headers.forEach(h => { persisted[h] = finalMapping[h] || ""; });
    window.MAPPING_CONFIG.saveOverrides(_pendingSource.sourceId, persisted);
  }

  // Validate
  validateSource(_pendingSource.sourceId);

  document.getElementById("mapperModal").classList.remove("show");
  _pendingSource = null;
  renderHub();
}

// ═══════════════════════════════════════════════════════════════════════════
// VALIDATION ENGINE
// ═══════════════════════════════════════════════════════════════════════════
function validateSource(sourceId){
  const s = SOURCES[sourceId];
  const meta = SOURCE_REGISTRY.find(m => m.id === sourceId);
  const issues = [];
  const mappedCanons = new Set(Object.values(s.mapping).filter(Boolean));

  // 1. Required canonical fields for this source
  (meta.required || []).forEach(req => {
    if (!mappedCanons.has(req)) {
      issues.push({severity:"CRITICAL", source:sourceId,
        msg:`Required field "${req}" is not mapped to any column.`});
    }
  });

  // 2. Duplicate-mapping detection: two headers mapped to the SAME canonical field
  //    (the second-best wins in extract; flag so admin knows)
  const canonToHeaders = {};
  Object.entries(s.mapping).forEach(([h, canon]) => {
    if (!canon) return;
    (canonToHeaders[canon] = canonToHeaders[canon] || []).push(h);
  });
  Object.entries(canonToHeaders).forEach(([canon, headers]) => {
    if (headers.length > 1) {
      issues.push({severity:"WARN", source:sourceId,
        msg:`Duplicate mapping: columns ${headers.map(h=>`"${h}"`).join(", ")} all map to "${canon}". First non-empty value wins.`});
    }
  });

  // 3. Extra / unknown columns: not a problem, but surface as INFO so admin sees they're being ignored
  const unmappedHeaders = s.headers.filter(h => !s.mapping[h]);
  if (unmappedHeaders.length) {
    issues.push({severity:"INFO", source:sourceId,
      msg:`${unmappedHeaders.length} column${unmappedHeaders.length>1?"s":""} ignored (no canonical mapping): ${unmappedHeaders.slice(0,5).map(h=>`"${h}"`).join(", ")}${unmappedHeaders.length>5?", …":""}`});
  }

  // 4. Row-level: missing employee IDs, duplicates, type validation, custom validators
  const nikSeen = {};
  const nameSeen = {};
  // Pre-compute per-header type/validator info for speed
  const headerChecks = {};
  for (const header in s.mapping) {
    const canon = s.mapping[header];
    const def = canon && FIELD_REGISTRY[canon];
    if (!def) continue;
    headerChecks[header] = {
      canon,
      typeFn: def.type ? TYPE_VALIDATORS[def.type] : null,
      typeName: def.type,
      customFn: def.validator ? VALIDATORS[def.validator] : null
    };
  }
  // Cap type-error spam: stop after 25 per canonical field
  const typeErrCount = {};
  const TYPE_ERR_CAP = 25;

  s.raw.forEach((row, i) => {
    const rn = i + 2;  // Excel row number (1=header)
    const nik = _extractCanonValue(row, s.mapping, "NIK");
    const name = _extractCanonValue(row, s.mapping, "Employee Name");

    if (mappedCanons.has("NIK") && !nik) {
      issues.push({severity:"WARN", source:sourceId, msg:`Row ${rn}: missing NIK.`});
    }
    if (mappedCanons.has("Employee Name") && !name) {
      issues.push({severity:"WARN", source:sourceId, msg:`Row ${rn}: missing Employee Name.`});
    }

    // Duplicate-row detection (by NIK preferred, name as fallback)
    if (nik) {
      const k = _norm(nik);
      if (nikSeen[k]) issues.push({severity:"WARN", source:sourceId,
        msg:`Row ${rn}: duplicate NIK "${nik}" (first seen row ${nikSeen[k]}).`});
      else nikSeen[k] = rn;
    } else if (name) {
      const k = _norm(name);
      if (nameSeen[k]) issues.push({severity:"INFO", source:sourceId,
        msg:`Row ${rn}: duplicate Employee Name "${name}" — and no NIK to disambiguate.`});
      else nameSeen[k] = rn;
    }

    // Cell-level: data-type + custom validators
    for (const header in headerChecks) {
      const chk = headerChecks[header];
      const v = row[header];
      // Type check (skip empty cells — empty is always OK; missing-required is checked separately)
      if (chk.typeFn && !chk.typeFn(v)) {
        typeErrCount[chk.canon] = (typeErrCount[chk.canon] || 0) + 1;
        if (typeErrCount[chk.canon] <= TYPE_ERR_CAP) {
          issues.push({severity:"WARN", source:sourceId,
            msg:`Row ${rn}: invalid data type for ${chk.canon} (expected ${chk.typeName}): "${v}".`});
        } else if (typeErrCount[chk.canon] === TYPE_ERR_CAP + 1) {
          issues.push({severity:"INFO", source:sourceId,
            msg:`(More ${chk.canon} type errors suppressed — only first ${TYPE_ERR_CAP} shown.)`});
        }
      }
      // Custom validator (e.g. readiness, flightrisk allowed values)
      if (chk.customFn && !chk.customFn(v)) {
        issues.push({severity:"WARN", source:sourceId,
          msg:`Row ${rn}: invalid ${chk.canon} value "${v}".`});
      }
    }
  });

  s.issues = issues;
  const hasCritical = issues.some(i => i.severity === "CRITICAL");
  s.status = hasCritical ? "blocked" : (issues.length ? "warn" : "ready");
}

// ═══════════════════════════════════════════════════════════════════════════
// CLEAR / REMOVE SOURCE
// ═══════════════════════════════════════════════════════════════════════════
function clearSource(sourceId){
  const s = SOURCES[sourceId];
  s.file = null; s.raw = []; s.headers = []; s.mapping = {}; s.issues = []; s.status = "empty";
  renderHub();
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function _extractCanonValue(row, mapping, canonical){
  for (const h in mapping) {
    if (mapping[h] === canonical) {
      const v = row[h];
      if (v !== undefined && v !== null && String(v).trim() !== "" && String(v).trim().toLowerCase() !== "nan") {
        return String(v).trim();
      }
    }
  }
  return "";
}
function _esc(s){return String(s||"").replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);}

function _alert(title, sub, rows){
  document.getElementById("vTitle").textContent = title;
  document.getElementById("vSub").textContent = sub || "";
  const em = document.getElementById("vErrors"); em.innerHTML = "";
  (rows||[]).forEach(r => {
    const d = document.createElement("div");
    d.className = "modal-err-row";
    d.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg><span class="modal-err-txt">${_esc(r.msg)}</span>`;
    em.appendChild(d);
  });
  document.getElementById("vOk").style.display = "none";
  document.getElementById("validModal").classList.add("show");
}
function closeModal(){document.getElementById("validModal").classList.remove("show");}
function confirmLoad(){closeModal();}  // legacy stub — not used by new flow

// ═══════════════════════════════════════════════════════════════════════════
// ISSUES PANEL
// ═══════════════════════════════════════════════════════════════════════════
let _issuesActiveTab = "all";
function openIssuesPanel(){
  _issuesActiveTab = "all";
  _renderIssuesPanel();
  document.getElementById("issuesModal").classList.add("show");
}
function closeIssuesPanel(){document.getElementById("issuesModal").classList.remove("show");}
function _renderIssuesPanel(){
  const all = Object.values(SOURCES).flatMap(s => s.issues);
  const tabs = document.getElementById("issuesTabs");
  const counts = {all: all.length};
  SOURCE_REGISTRY.forEach(m => counts[m.id] = SOURCES[m.id].issues.length);
  tabs.innerHTML = `<button class="issues-tab ${_issuesActiveTab==='all'?'active':''}" onclick="_setIssuesTab('all')">All <span class="cnt">${counts.all}</span></button>` +
    SOURCE_REGISTRY.filter(m => counts[m.id]).map(m =>
      `<button class="issues-tab ${_issuesActiveTab===m.id?'active':''}" onclick="_setIssuesTab('${m.id}')">${m.label} <span class="cnt">${counts[m.id]}</span></button>`
    ).join("");

  const list = document.getElementById("issuesList");
  const filtered = _issuesActiveTab === "all" ? all : SOURCES[_issuesActiveTab].issues;
  if (!filtered.length) { list.innerHTML = `<div class="iss-empty">No issues to show.</div>`; return; }
  // Sort: CRITICAL > ERROR > WARN > INFO
  const sevRank = {CRITICAL:0, ERROR:1, WARN:2, INFO:3};
  const sorted = [...filtered].sort((a,b) => (sevRank[a.severity]||9) - (sevRank[b.severity]||9));
  list.innerHTML = sorted.slice(0, 500).map(iss => {
    const srcLabel = (SOURCE_REGISTRY.find(m => m.id === iss.source)||{}).label || iss.source;
    return `<div class="iss-row">
      <span class="iss-sev sev-${iss.severity}">${iss.severity}</span>
      <span class="iss-src">${_esc(srcLabel)}</span>
      <span class="iss-msg">${_esc(iss.msg)}</span>
    </div>`;
  }).join("") + (sorted.length > 500 ? `<div class="iss-empty">…${sorted.length-500} more issues hidden.</div>` : "");
}
function _setIssuesTab(t){_issuesActiveTab=t;_renderIssuesPanel();}

// ═══════════════════════════════════════════════════════════════════════════
// DATA INTEGRATION LAYER
// ═══════════════════════════════════════════════════════════════════════════
// Strategy:
//   - For the PRIMARY source (SAP / first loaded with rows), each row becomes
//     a "base" record. If SAP has a row-per-position (succession-style data),
//     each row is preserved (since the dashboard works at row granularity).
//   - For each base row, we look up the same employee in OTHER sources by:
//       1. NIK (primary key)
//       2. Employee Name (fallback)
//     and overlay any canonical fields that the base row doesn't already have
//     (i.e., enrichment, not overwrite).
//   - The output is a flat list of merged rows keyed by canonical field names —
//     exactly the shape the existing dashboard's g(row,...) helper consumes.
//
// Result: the dashboard treats this multi-source merged set identically to
// the original single-file payload.
//
function buildIntegratedData(){
  const loadedSources = SOURCE_REGISTRY
    .map(m => ({meta:m, src:SOURCES[m.id]}))
    .filter(x => x.src.raw.length > 0);

  if (!loadedSources.length) return [];

  // Pick primary: prefer SAP if loaded, else first loaded with most rows
  let primary = loadedSources.find(x => x.meta.id === "sap");
  if (!primary) primary = loadedSources.slice().sort((a,b) => b.src.raw.length - a.src.raw.length)[0];

  // Build per-source indices for fast lookup
  const otherSources = loadedSources.filter(x => x.meta.id !== primary.meta.id);
  const indices = otherSources.map(x => {
    const byNik = {}, byName = {};
    x.src.raw.forEach(row => {
      const nik = _extractCanonValue(row, x.src.mapping, "NIK");
      const name = _extractCanonValue(row, x.src.mapping, "Employee Name");
      if (nik) {
        const k = _norm(nik);
        if (!byNik[k]) byNik[k] = [];
        byNik[k].push(row);
      }
      if (name) {
        const k = _norm(name);
        if (!byName[k]) byName[k] = [];
        byName[k].push(row);
      }
    });
    return {meta:x.meta, src:x.src, byNik, byName};
  });

  // Build merged output
  const merged = [];
  primary.src.raw.forEach(primRow => {
    // Convert primary row into canonical row first
    const canonRow = _rowToCanonical(primRow, primary.src.mapping);
    // Tag provenance for debugging / future API integration
    canonRow._sources = [primary.meta.id];

    const nik = canonRow["NIK"] || "";
    const name = canonRow["Employee Name"] || "";
    const nikKey = nik ? _norm(nik) : null;
    const nameKey = name ? _norm(name) : null;

    // Enrich from each other source
    indices.forEach(ix => {
      let matches = null;
      if (nikKey && ix.byNik[nikKey]) matches = ix.byNik[nikKey];
      else if (nameKey && ix.byName[nameKey]) matches = ix.byName[nameKey];
      if (!matches || !matches.length) return;

      // If multiple matches in the other source (e.g., PCD has many training rows),
      // collapse them: take the first non-empty value for scalar canonical fields,
      // concatenate for list-like fields (training/certs/projects).
      const LIST_FIELDS = new Set(["Training Name","Training Plan","Certification Plan","Project Assignment"]);
      const collected = {};
      matches.forEach(m => {
        const cm = _rowToCanonical(m, ix.src.mapping);
        for (const k in cm) {
          if (!cm[k]) continue;
          if (LIST_FIELDS.has(k)) {
            collected[k] = collected[k] ? collected[k] + "; " + cm[k] : cm[k];
          } else if (!collected[k]) {
            collected[k] = cm[k];
          }
        }
      });
      // Apply enrichment (do not overwrite primary's non-empty values)
      for (const k in collected) {
        if (!canonRow[k]) canonRow[k] = collected[k];
      }
      canonRow._sources.push(ix.meta.id);
    });

    merged.push(canonRow);
  });

  return merged;
}

function _rowToCanonical(row, mapping){
  const out = {};
  for (const header in mapping) {
    const canon = mapping[header];
    if (!canon) continue;
    const v = row[header];
    if (v === undefined || v === null) continue;
    const sv = String(v).trim();
    if (sv === "" || sv.toLowerCase() === "nan") continue;
    out[canon] = sv;
  }
  return out;
}

// ═══════════════════════════════════════════════════════════════════════════
// BUILD & ENTER DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function buildAndEnter(){
  // Cross-source validation: warn about employees in TM/PCD but not in SAP
  _runCrossSourceValidation();

  const merged = buildIntegratedData();
  if (!merged.length) {
    _alert("Cannot Build", "No employee data could be integrated.", [{severity:"ERROR", msg:"At least one source must contain rows mapped to NIK or Employee Name."}]);
    return;
  }

  ALL = merged;

  // Update topbar
  const loadedSrcs = SOURCE_REGISTRY.filter(m => SOURCES[m.id].raw.length).map(m => m.label);
  document.getElementById("topFile").textContent = loadedSrcs.length === 1 ? loadedSrcs[0] : `${loadedSrcs.length} sources merged`;
  document.getElementById("topRec").textContent = ALL.length + " rows";

  document.getElementById("upload-screen").style.display = "none";
  document.getElementById("dashboard").style.display = "block";

  buildFilters();
  FILTERED = [...ALL];
  renderAll();
}

function _runCrossSourceValidation(){
  // Build NIK universe from primary
  const primaryMeta = SOURCES["sap"].raw.length ? SOURCE_REGISTRY[0] : null;
  if (!primaryMeta) return;
  const primary = SOURCES[primaryMeta.id];
  const primNiks = new Set();
  primary.raw.forEach(r => {
    const nik = _extractCanonValue(r, primary.mapping, "NIK");
    if (nik) primNiks.add(_norm(nik));
  });
  if (!primNiks.size) return;

  // Strip prior cross-source notices to avoid duplication across rebuilds
  Object.values(SOURCES).forEach(s => {
    s.issues = s.issues.filter(i => !i.crossSource);
  });

  SOURCE_REGISTRY.slice(1).forEach(m => {
    const s = SOURCES[m.id];
    if (!s.raw.length) return;
    let orphans = 0;
    s.raw.forEach((row,i) => {
      const nik = _extractCanonValue(row, s.mapping, "NIK");
      if (nik && !primNiks.has(_norm(nik))) orphans++;
    });
    if (orphans) {
      s.issues.push({severity:"INFO", source:m.id, crossSource:true,
        msg:`${orphans} employee${orphans>1?"s":""} in this source not found in SAP master.`});
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// RESET
// ═══════════════════════════════════════════════════════════════════════════
function resetApp(){
  document.getElementById("upload-screen").style.display = "flex";
  document.getElementById("dashboard").style.display = "none";
  document.getElementById("fileInput").value = "";
  Object.values(CHARTS).forEach(c => { try { c.destroy(); } catch(e){} });
  CHARTS = {}; ALL = []; FILTERED = [];
  SOURCE_REGISTRY.forEach(m => {
    const s = SOURCES[m.id];
    s.file = null; s.raw = []; s.headers = []; s.mapping = {}; s.issues = []; s.status = "empty";
  });
  renderHub();
}

// init moved to useEffect in App.jsx

function g(row,...keys){for(const k of keys){const v=row[k];if(v!==undefined&&v!==null&&v!==""&&String(v).trim()!=="nan")return String(v).trim();}return "";}
function isYes(v){const s=String(v||"").toUpperCase().trim();return s==="YES"||s==="Y"||s==="1"||s==="TRUE";}
function calcAge(dob){if(!dob)return null;const d=new Date(dob);if(isNaN(d))return null;return Math.floor((Date.now()-d.getTime())/31557600000);}
function ageFromGroup(grp){
  if(!grp)return null;
  const s=grp.toString().trim();
  if(s==="< 30"||s==="<30")return 25;
  if(s==="> 50"||s===">50")return 52;
  if(s==="51-55")return 53;
  if(s==="> 55"||s===">55")return 57;
  const m=s.match(/(\d+)/);
  return m?parseInt(m[1])+2:null;
}
function retRisk(age){if(age===null||age===undefined||age==="")return"Unknown";if(age>=51)return"High";if(age>=41)return"Moderate";return"Low";}
function retRiskFromGroup(grp){if(!grp||grp.toString().trim()==="")return"Unknown";return retRisk(ageFromGroup(grp));}
function rColor(txt){const t=(txt||"").toLowerCase();if(t.includes("now"))return"b-green";if(t.includes("1-3")||t.includes("later"))return"b-amber";if(t.includes("3-5")||t.includes("future"))return"b-blue";return"b-gray";}
function countBy(arr,...keys){const m={};arr.forEach(d=>{const v=g(d,...keys)||"Unknown";m[v]=(m[v]||0)+1;});return m;}
function sorted(m,n=999){return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,n);}

function buildPM(data){
  const m={};
  data.forEach(row=>{
    const pid=g(row,"Position ID");if(!pid)return;
    if(!m[pid]){
      const grp=g(row,"Group Age");
      const age=ageFromGroup(grp);
      // collect position technical skills (up to 5)
      const techSkills=[];
      for(let i=1;i<=5;i++){const desc=g(row,"Technical Skill Required "+i+" Desc");const thr=g(row,"Technical Skill "+i+" Treshold Desc");if(desc)techSkills.push({desc,threshold:thr});}
      // collect position soft skills (up to 8)
      const softSkills=[];
      for(let i=1;i<=8;i++){
        const desc=g(row,"Soft Skill Required "+i+" Desc")||(i===2?g(row,"Char255"):"");
        const thr=g(row,"Soft Skill Required "+i+" Treshold Desc");
        if(desc)softSkills.push({desc,threshold:thr});
      }
      m[pid]={id:pid,title:g(row,"Position Name"),family:g(row,"Job Family Successor"),
        positionStatus:g(row,"Position Status"),
        pillarName:g(row,"Pillar Name"),
        level:g(row,"Current Level"),
        nik:g(row,"NIK"),
        startDate:g(row,"Start Date"),
        endDate:g(row,"End Date"),
        lastPromotion:g(row,"Last Promotion Date"),
        lastRotation:g(row,"Last Rotation Date"),
        incumbent:g(row,"Employee Name"),incAgeGrp:grp,incAge:age,
        retRisk:retRiskFromGroup(grp),
        flightRisk:g(row,"Incumbent Flight Risk")||"Unknown",
        clevel:isYes(g(row,"C-Level")),
        critical:isYes(g(row,"Critical Position")),
        bu:g(row,"Business Unit Name"),buId:g(row,"Business Unit ID"),
        group:g(row,"Group Name"),groupId:g(row,"Group ID"),
        directorate:g(row,"Directorate Name"),directorateId:g(row,"Directorate ID"),
        department:g(row,"Departement Name"),departmentId:g(row,"Departement ID"),
        division:g(row,"Division"),divisionId:g(row,"Division ID"),
        supervisor:g(row,"Supervisor Name"),supervisorId:g(row,"Supervisor ID"),
        spvPositionName:g(row,"SPV Position Name"),
        techSkills,softSkills,
        years:"",successors:[]};
    }
    const sn=g(row,"Successor Name");
    if(sn&&!m[pid].successors.find(s=>s.name===sn)){
      // collect successor tech skills (up to 5)
      const sTech=[];
      for(let i=1;i<=5;i++){const desc=g(row,"Successor Tech Skill "+i+" Desc");const thr=g(row,"Successor Tech Skill "+i+" Threshold Desc");if(desc)sTech.push({desc,threshold:thr});}
      // collect successor additional tech skills (up to 5)
      const sAddTech=[];
      for(let i=1;i<=5;i++){const desc=g(row,"Successor Additional Tech Skill "+i+" Desc");const thr=g(row,"Succ Add Tech Skill "+i+" Threshold Desc");if(desc)sAddTech.push({desc,threshold:thr});}
      // collect successor soft skills (up to 8)
      const sSoft=[];
      for(let i=1;i<=8;i++){const desc=g(row,"Successor Soft Skill "+i+" Desc");const thr=g(row,"Successor Soft Skill "+i+" Threshold Desc");if(desc)sSoft.push({desc,threshold:thr});}
      // collect training (3 slots with same column names)
      const trainings=[];
      const allRows=Object.keys(row);
      // Training Name appears 3 times; use raw row array approach via known col names
      const tNames=[g(row,"Training Name")].filter(Boolean);
      const tPlans=[g(row,"Training Plan")].filter(Boolean);
      m[pid].successors.push({
        name:sn,
        succEmp:g(row,"Succession Emp Number"),
        succPosition:g(row,"Successor Position"),
        succPositionId:g(row,"Successor Position ID"),
        readiness:g(row,"Final Score Readiness Text"),
        readinessScore:g(row,"Final Score Readiness"),
        asTalent:isYes(g(row,"Successor As Talent")),
        calibrationStatus:g(row,"Calibration Status"),
        level:g(row,"Level Successor"),
        family:g(row,"Job Family Successor"),
        bu:g(row,"Business Unit Successor"),
        flightRisk:g(row,"Successor Flight Risk"),
        jobComplexity:g(row,"Job Complexity"),
        yearsExp:g(row,"Years of Experience"),
        techKnowledge:g(row,"Technical Knowledge"),
        softKnowledge:g(row,"Softskill Knowledge"),
        coach:g(row,"Coach"),
        mentor:g(row,"Mentor"),
        mentorNom1:g(row,"Mentor Nomination 1 Name"),
        mentorNom2:g(row,"Mentor Nomination 2 Name"),
        projectAssignment:g(row,"Project Assignment"),
        certPlan:g(row,"Certification Plan"),
        trainingName:g(row,"Training Name"),
        trainingPlan:g(row,"Training Plan"),
        sTech,sAddTech,sSoft,
      });
    }
  });
  return m;
}

const FDEFS=[
  {id:"f_fam",label:"Job Family",keys:["Job Family Successor"]},
  {id:"f_bu",label:"Business Unit",keys:["Business Unit Name"]},
  {id:"f_group",label:"Group",keys:["Group Name"]},
  {id:"f_cl",label:"C-Level",keys:["C-Level"]},
  {id:"f_crit",label:"Critical",keys:["Critical Position"]},
  {id:"f_ready",label:"Readiness",keys:["Final Score Readiness Text"]},
  {id:"f_talent",label:"As Talent",keys:["As Talent Calibrated"]},
];
function buildFilters(){
  const bar=document.getElementById("filterBar");
  bar.innerHTML='<span class="filter-lbl">Filter by:</span>';
  FDEFS.forEach(fd=>{
    const vals=[...new Set(ALL.map(d=>g(d,...fd.keys)).filter(v=>v&&v!=="undefined"))].sort();
    if(!vals.length)return;
    const wrap=document.createElement("div");wrap.style.cssText="display:flex;align-items:center;gap:5px;";
    const lbl=document.createElement("span");lbl.textContent=fd.label+" :";lbl.style.cssText="font-size:11px;font-weight:600;color:var(--t2);white-space:nowrap;";
    const sel=document.createElement("select");sel.id=fd.id;
    sel.innerHTML='<option value="" disabled selected style="color:#aaa;">All</option>'+vals.map(v=>"<option>"+v+"</option>").join("");
    sel.onchange=applyFilters;
    wrap.appendChild(lbl);wrap.appendChild(sel);bar.appendChild(wrap);
  });
  const rb=document.createElement("button");rb.className="btn-rf";rb.textContent="Reset";
  rb.onclick=()=>{document.querySelectorAll(".filter-bar select").forEach(s=>s.value="");applyFilters();};
  bar.appendChild(rb);
  // Build tab-level BU filters
  buildTabBuFilter("overviewBuFilter",()=>{renderBUChart();renderAgeChart();renderLayerChart();renderJobFamilyChart();renderReadinessChart();renderFlightChart();renderCritChart();renderRetChart();});
  document.getElementById("detailClFilter").onchange=()=>renderDetail();
  buildTabBuFilter("posBuFilter",()=>renderPositions());
  buildTabBuFilter("detailBuFilter",()=>renderDetail());
}
function applyOverviewUniqFilter(){
  // Re-render all overview charts with the unique filter applied
  renderBUChart();renderAgeChart();renderLayerChart();renderJobFamilyChart();
  renderReadinessChart();renderFlightChart();renderCritChart();renderRetChart();
}
function getOverviewData(){
  const src=tabBuFilter("overviewBuFilter","overviewClFilter");
  const mode=(document.getElementById("overviewUniqFilter")||{}).value||"all";
  if(mode==="all")return src;
  // Match snap modal semantics: count occurrences of Successor Name across the
  // current (BU-filtered) source. Unique = name appears exactly once; Non-Unique
  // = name appears more than once. Rows with no Successor Name are excluded.
  const count={};
  src.forEach(r=>{const n=(g(r,"Successor Name")||"").toLowerCase().trim();if(n)count[n]=(count[n]||0)+1;});
  if(mode==="unique")
    return src.filter(r=>{const n=(g(r,"Successor Name")||"").toLowerCase().trim();return n&&count[n]===1;});
  // nonunique
  return src.filter(r=>{const n=(g(r,"Successor Name")||"").toLowerCase().trim();return n&&count[n]>1;});
}
function applyFilters(){
  FILTERED=ALL.filter(d=>FDEFS.every(fd=>{
    const el=document.getElementById(fd.id);if(!el||!el.value)return true;
    return g(d,...fd.keys)===el.value;
  }));
  document.getElementById("topRec").textContent=FILTERED.length+" rows";
  DETAIL_PAGE=1;POS_PAGE=1;
  renderAll();
}
function switchTab(btn,name){
  document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));
  document.getElementById("tab-"+name).classList.add("active");btn.classList.add("active");
}
// Build successor load map: succName -> [{posTitle, readiness}]
let SUCC_LOAD_MAP={};
function buildSuccLoadMap(){
  SUCC_LOAD_MAP={};
  const pm=buildPM(FILTERED);
  Object.values(pm).forEach(p=>{
    p.successors.forEach(s=>{
      if(!SUCC_LOAD_MAP[s.name])SUCC_LOAD_MAP[s.name]=[];
      SUCC_LOAD_MAP[s.name].push({pos:p.title,readiness:s.readiness,clevel:p.clevel,critical:p.critical});
    });
  });
}
function openSuccPopup(name,event){
  event.stopPropagation();
  const positions=SUCC_LOAD_MAP[name]||[];
  document.getElementById("succPopupName").textContent=name;
  document.getElementById("succPopupSub").textContent="Successor to "+positions.length+" position"+(positions.length!==1?"s":"");
  document.getElementById("succPopupBody").innerHTML=positions.map(p=>`
    <div class="succ-pos-item">
      <div>
        <div style="font-weight:600;color:var(--navy)">${p.pos}</div>
        <div style="display:flex;gap:4px;margin-top:3px;">
          ${p.clevel?'<span class="b b-gold" style="font-size:9px">C-Level</span>':""}
          ${p.critical?'<span class="b b-red" style="font-size:9px">Critical</span>':""}
        </div>
      </div>
      ${p.readiness?'<span class="b '+rColor(p.readiness)+'" style="font-size:9px">'+p.readiness+'</span>':""}
    </div>`).join("")||'<div style="color:var(--t3);font-size:11px;padding:6px 0">No positions found.</div>';
  document.getElementById("succPopupBackdrop").classList.add("show");
  document.body.style.overflow="hidden";
}
function closeSuccPopup(){
  document.getElementById("succPopupBackdrop").classList.remove("show");
  document.body.style.overflow="";
}
document.addEventListener("keydown",e=>{if(e.key==="Escape")closeSuccPopup();});


// ── TAB BU FILTERS ──
function buildTabBuFilter(selectId, onChangeFn){
  const bus=[...new Set(ALL.map(r=>g(r,"Business Unit Name")).filter(v=>v))].sort();
  const sel=document.getElementById(selectId);if(!sel)return;
  const cur=sel.value;
  sel.innerHTML='<option value="">All Business Units</option>'+bus.map(v=>'<option'+(v===cur?' selected':'')+'>'+v+'</option>').join('');
  sel.onchange=onChangeFn;
}
function getTabBu(selectId){return(document.getElementById(selectId)||{}).value||"";}

function renderAll(){
  renderSnap();renderPosture();renderCLevelPosture();renderReadinessPosture();renderBench();
  renderBUChart();renderAgeChart();renderLayerChart();renderJobFamilyChart();
  renderReadinessChart();renderFlightChart();renderCritChart();renderRetChart();
  renderPositions();renderSimPanel();renderDetail();renderExecSummary();
}

function renderSnap(){
  const pm = buildPM(FILTERED);
  const arr = Object.values(pm);

  const allEmpNames = FILTERED.map(r=>g(r,"Employee Name"));
  const empUnknown = allEmpNames.filter(v=>!v||v.toString().trim()==="").length;
  const empKnown   = allEmpNames.filter(v=>v&&v.toString().trim()!=="").map(v=>v.toString().trim().toLowerCase().replace(/\s+/g," "));
  const empCount={};empKnown.forEach(n=>{empCount[n]=(empCount[n]||0)+1;});
  const empDupes   = empKnown.filter(n=>empCount[n]>1).length;  // rows where name appears >1
  const totalEmp   = empKnown.filter(n=>empCount[n]===1).length; // rows where name appears exactly once

const talentSet = new Set();
FILTERED.forEach(r=>{
  const val=(g(r,"As Talent Calibrated")||"").toString().trim().toUpperCase();
  if(val==="YES"){const name=g(r,"Employee Name");if(name)talentSet.add(name);}
});
// totalTalent now calculated below from talentKnownRows
// talent unknown = rows where As Talent Calibrated = YES but Employee Name is blank
const talentUnknown=FILTERED.filter(r=>isYes(g(r,"As Talent Calibrated"))&&!g(r,"Employee Name")).length;
// talent dupes = all rows with a name that appears more than once
const talentKnownRows=FILTERED.filter(r=>isYes(g(r,"As Talent Calibrated"))&&g(r,"Employee Name")).map(r=>g(r,"Employee Name").toString().trim().toLowerCase().replace(/\s+/g," "));
const talentCount={};talentKnownRows.forEach(n=>{talentCount[n]=(talentCount[n]||0)+1;});
const talentDupes=talentKnownRows.filter(n=>talentCount[n]>1).length;
const totalTalent=talentKnownRows.filter(n=>talentCount[n]===1).length;

  // 🔧 Normalize function (ONLY ONCE)
  function normalizeName(name){
    return (name||"").toLowerCase().trim().replace(/\s+/g," ");
  }

  // 👥 Successors (unique, cleaned)
  const allSuccRaw=FILTERED.map(r=>g(r,"Successor Name"));
  const succUnknown=allSuccRaw.filter(v=>!v||v.toString().trim()==="").length;
  const succKnown=allSuccRaw.filter(v=>v&&v.toString().trim()!=="");
  const successorSet=new Set(succKnown.map(normalizeName));
  const succCount={};succKnown.forEach(n=>{succCount[normalizeName(n)]=(succCount[normalizeName(n)]||0)+1;});
  const succDupes=succKnown.filter(n=>succCount[normalizeName(n)]>1).length;
  const totalSuccessor=succKnown.filter(n=>succCount[normalizeName(n)]===1).length;

  // With/Without Successor: employee HAS a successor if their row contains a Successor Name
  const _empHasSuccSet=new Set(
    FILTERED.filter(r=>g(r,"Employee Name")&&g(r,"Successor Name"))
            .map(r=>g(r,"Employee Name").toString().trim().toLowerCase())
  );
  const empWithSuccCount=empKnown.filter(n=>empCount[n]===1&&_empHasSuccSet.has(n)).length;
  const empWithoutSuccCount=totalEmp-empWithSuccCount;
  const empNonUniqWithSucc=empKnown.filter(n=>empCount[n]>1&&_empHasSuccSet.has(n)).length;
  const empNonUniqWithoutSucc=empKnown.filter(n=>empCount[n]>1&&!_empHasSuccSet.has(n)).length;
  const talentWithSuccCount=talentKnownRows.filter(n=>talentCount[n]===1&&_empHasSuccSet.has(n)).length;
  const talentWithoutSuccCount=totalTalent-talentWithSuccCount;
  const talentNonUniqWithSucc=talentKnownRows.filter(n=>talentCount[n]>1&&_empHasSuccSet.has(n)).length;
  const talentNonUniqWithoutSucc=talentKnownRows.filter(n=>talentCount[n]>1&&!_empHasSuccSet.has(n)).length;

  // 📊 Percentages
  const empPct = totalEmp>0 ? (totalTalent/totalEmp*100).toFixed(1) : 0;
  const succPct = totalEmp>0 ? (totalSuccessor/totalEmp*100).toFixed(1) : 0;

  // 💎 Hidden Gems: successors to Critical Positions who are NOT themselves talent-calibrated
  const talentIncNames=new Set(FILTERED.filter(r=>isYes(g(r,"As Talent Calibrated"))).map(r=>g(r,"Employee Name")).filter(v=>v));
  const hiddenGemsSet=new Set(
    FILTERED.filter(r=>isYes(g(r,"Critical Position"))&&g(r,"Successor Name")&&!talentIncNames.has(g(r,"Successor Name")))
            .map(r=>g(r,"Successor Name"))
  );
  const hiddenGems=hiddenGemsSet.size;
  const hgPct=totalEmp>0?(hiddenGems/totalEmp*100).toFixed(1):"0";

  // 🎯 SNAP CARDS
  const snaps=[
    {key:"employee",label:"EMPLOYEE",val:totalEmp,pct:"",empUnknown,empDupes,withSucc:empWithSuccCount,withoutSucc:empWithoutSuccCount,nonUniqWithSucc:empNonUniqWithSucc,nonUniqWithoutSucc:empNonUniqWithoutSucc,totalWithSucc:empWithSuccCount+empNonUniqWithSucc,totalWithoutSucc:empWithoutSuccCount+empNonUniqWithoutSucc,
     rawTotal: allEmpNames.length,
     icon:'<path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/>',
     iconBg:"rgba(30,111,217,.22)",iconStroke:"#60A5FA"},

    {key:"talent",label:"TALENT",val:totalTalent,pct:empPct+"%",unknown:talentUnknown,dupes:talentDupes,withSucc:talentWithSuccCount,withoutSucc:talentWithoutSuccCount,nonUniqWithSucc:talentNonUniqWithSucc,nonUniqWithoutSucc:talentNonUniqWithoutSucc,totalWithSucc:talentWithSuccCount+talentNonUniqWithSucc,totalWithoutSucc:talentWithoutSuccCount+talentNonUniqWithoutSucc,
     icon:'<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
     iconBg:"rgba(124,58,237,.22)",iconStroke:"#A78BFA"},

    {key:"successor",label:"SUCCESSOR",val:totalSuccessor,pct:succPct+"%",unknown:succUnknown,dupes:succDupes,
     icon:'<path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>',
     iconBg:"rgba(16,185,129,.22)",iconStroke:"#34D399"},
  ];

  document.getElementById("snapGrid").innerHTML = snaps.map(s=>`
    <div class="snap-card" id="snap-card-${s.key}" onclick="openSnapModal('${s.key}')" title="Click to expand details below">
      <div class="snap-body" style="width:100%">
        <div class="snap-label">${s.label}</div>
        ${(()=>{
          const unk=s.key==='employee'?s.empUnknown:(s.unknown||0);
          const dup=s.key==='employee'?s.empDupes:(s.dupes||0);
          const total=s.key==='employee'?(s.val+dup+unk-dup+dup):s.val+dup; // raw row count
          // For employee: total rows = unique rows + non-unique rows - overlap... simplest: empKnown.length + empUnknown
          const displayTotal=s.key==='employee'?s.rawTotal:(s.val+dup);
          return '<div style="display:flex;align-items:center;gap:10px;margin-bottom:2px;">'
          +'<div class="snap-icon" style="background:'+s.iconBg+';width:38px;height:38px;border-radius:9px;flex-shrink:0;">'
          +'<svg viewBox="0 0 24 24" fill="none" stroke="'+s.iconStroke+'" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" style="width:18px;height:18px;">'+s.icon+'</svg>'
          +'</div>'
          +'<div>'
          +'<div class="snap-val">'+displayTotal.toLocaleString()+'</div>'
          +''
          +'</div>'
          +'</div>'
          +'<div style="margin-top:8px;display:flex;flex-direction:column;gap:4px;border-top:1px solid rgba(255,255,255,.1);padding-top:8px;">'
          +'<div style="font-size:12px;color:rgba(255,255,255,.5);display:flex;justify-content:space-between;gap:8px;"><span>Unique</span><span style="font-weight:700;color:rgba(255,255,255,.95)">'+s.val.toLocaleString()+'</span></div>'
          +((s.key==='talent'||s.key==='employee')?'<div style="font-size:11px;color:rgba(255,255,255,.38);display:flex;justify-content:space-between;gap:8px;padding-left:12px;"><span>With Successor</span><span style="font-weight:600;color:rgba(255,255,255,.6)">'+(s.withSucc||0).toLocaleString()+'</span></div>'+'<div style="font-size:11px;color:rgba(255,255,255,.38);display:flex;justify-content:space-between;gap:8px;padding-left:12px;"><span>Without Successor</span><span style="font-weight:600;color:rgba(255,255,255,.6)">'+(s.withoutSucc||0).toLocaleString()+'</span></div>':'')
          +'<div style="font-size:12px;color:rgba(255,255,255,.5);display:flex;justify-content:space-between;gap:8px;"><span>Non-Unique</span><span style="font-weight:700;color:rgba(255,255,255,.75)">'+dup.toLocaleString()+'</span></div>'
          +((s.key==='talent'||s.key==='employee')?'<div style="font-size:11px;color:rgba(255,255,255,.38);display:flex;justify-content:space-between;gap:8px;padding-left:12px;"><span>With Successor</span><span style="font-weight:600;color:rgba(255,255,255,.6)">'+(s.nonUniqWithSucc||0).toLocaleString()+'</span></div>'+'<div style="font-size:11px;color:rgba(255,255,255,.38);display:flex;justify-content:space-between;gap:8px;padding-left:12px;"><span>Without Successor</span><span style="font-weight:600;color:rgba(255,255,255,.6)">'+(s.nonUniqWithoutSucc||0).toLocaleString()+'</span></div>':'')
          +(s.key==='employee'||s.key==='talent'?'':'<div style="font-size:12px;color:rgba(255,255,255,.5);display:flex;justify-content:space-between;gap:8px;"><span>Unknown</span><span style="font-weight:700;color:rgba(255,255,255,.75)">'+unk.toLocaleString()+'</span></div>')
          +((s.key==='talent'||s.key==='employee')?
            '<div style="margin-top:6px;padding-top:6px;border-top:1px solid rgba(255,255,255,.15);display:flex;justify-content:space-between;gap:8px;font-size:12px;color:rgba(255,255,255,.65);"><span>Total With Successor</span><span style="font-weight:700;color:rgba(255,255,255,.95)">'+(s.totalWithSucc||0).toLocaleString()+'</span></div>'+
            '<div style="display:flex;justify-content:space-between;gap:8px;font-size:12px;color:rgba(255,255,255,.65);"><span>Total Without Successor</span><span style="font-weight:700;color:rgba(255,255,255,.95)">'+(s.totalWithoutSucc||0).toLocaleString()+'</span></div>'
          :'')
          +'</div>';
        })()}
        <div class="snap-click-hint">
          <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round">
            <circle cx="11" cy="11" r="8"/>
            <line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg> View details
        </div>
      </div>
    </div>
  `).join("");
}

function hexToRgb(hex){
  const r=parseInt(hex.slice(1,3),16),g2=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);
  return r+","+g2+","+b;
}

// ── SNAP MODAL ──
let SNAP_MODAL_DATA=[];
function openSnapModal(key){
  const pm=buildPM(FILTERED);
  const arr=Object.values(pm);
  let rows=[], allRows=[];
  let title="",sub="";
  SNAP_KEY=key;

  if(key==="employee"){
    title="Employee";sub="All incumbents on record";
    const seen=new Set();
    const _empHasSuccModal=new Set(FILTERED.filter(r=>g(r,"Employee Name")&&g(r,"Successor Name")).map(r=>g(r,"Employee Name").toString().trim().toLowerCase()));
    FILTERED.forEach(r=>{
      const n=g(r,"Employee Name");
      const p=arr.find(x=>x.incumbent===n)||{};
      const hasSuccessor=n?_empHasSuccModal.has(n.toString().trim().toLowerCase()):false;
      // Each row carries only its own successor — no cross-row aggregation
      const successorNames=g(r,"Successor Name")||"";
      const successorReadiness=g(r,"Final Score Readiness Text")||"";
      const rowObj={name:n||"—",role:g(r,"Position Name"),family:g(r,"Job Family Successor"),level:g(r,"Current Level"),
        bu:g(r,"Business Unit Name"),flightRisk:g(r,"Incumbent Flight Risk"),retRisk:p.retRisk||"",years:g(r,"Division"),
        clevel:isYes(g(r,"C-Level"))?"Yes":"",critical:isYes(g(r,"Critical Position"))?"Yes":"",hasSuccessor,successorNames,successorReadiness};
      allRows.push(rowObj);
      if(!n||seen.has(n))return; seen.add(n); rows.push(rowObj);
    });
  } else if(key==="talent"){
    title="Talent";sub="";
    const seen=new Set();
    const _talHasSuccModal=new Set(FILTERED.filter(r=>g(r,"Employee Name")&&g(r,"Successor Name")).map(r=>g(r,"Employee Name").toString().trim().toLowerCase()));
    FILTERED.forEach(r=>{
      if(!isYes(g(r,"As Talent Calibrated")))return;
      const n=g(r,"Employee Name");
      const hasSuccessor=n?_talHasSuccModal.has(n.toString().trim().toLowerCase()):false;
      // Each row carries only its own successor — no cross-row aggregation
      const successorNames=g(r,"Successor Name")||"";
      const successorReadiness=g(r,"Final Score Readiness Text")||"";
      const rowObj={name:n||"—",role:g(r,"Position Name"),family:g(r,"Job Family Successor"),level:g(r,"Current Level"),
        bu:g(r,"Business Unit Name"),flightRisk:g(r,"Incumbent Flight Risk"),clevel:isYes(g(r,"C-Level"))?"Yes":"",critical:isYes(g(r,"Critical Position"))?"Yes":"",hasSuccessor,successorNames,successorReadiness};
      allRows.push(rowObj);
      if(!n||seen.has(n))return; seen.add(n); rows.push(rowObj);
    });
  } else if(key==="successor_as_talent"){
    title="Successor as Talent";sub="Successors flagged as talent";
    const seen=new Set();
    FILTERED.filter(r=>isYes(g(r,"Successor As Talent"))&&g(r,"Successor Name")).forEach(r=>{
      const n=g(r,"Successor Name");
      const rowObj={name:n||"—",role:g(r,"Position Name"),family:g(r,"Job Family Successor"),
        level:g(r,"Level Successor"),bu:g(r,"Business Unit Successor"),
        readiness:g(r,"Final Score Readiness Text"),flightRisk:g(r,"Successor Flight Risk"),clevel:isYes(g(r,"C-Level"))?"Yes":""};
      allRows.push(rowObj);
      if(!n||seen.has(n))return; seen.add(n); rows.push(rowObj);
    });
  } else if(key==="new_potential_talent"){
    title="New Potential Talent";sub="Successors not yet flagged as talent";
    const seen=new Set();
    FILTERED.filter(r=>g(r,"Successor Name")&&!isYes(g(r,"As Talent Calibrated"))).forEach(r=>{
      const n=g(r,"Successor Name");
      const rowObj={name:n||"—",role:g(r,"Position Name"),family:g(r,"Job Family Successor"),
        level:g(r,"Level Successor"),bu:g(r,"Business Unit Successor"),
        readiness:g(r,"Final Score Readiness Text"),flightRisk:g(r,"Successor Flight Risk"),clevel:isYes(g(r,"C-Level"))?"Yes":""};
      allRows.push(rowObj);
      if(!n||seen.has(n))return; seen.add(n); rows.push(rowObj);
    });
  } else if(key==="hiddengems"){
    title="Hidden Gems";sub="Successors to Critical Positions who are not Talent Calibrated";
    const talentNames=new Set(FILTERED.filter(r=>isYes(g(r,"As Talent Calibrated"))).map(r=>g(r,"Employee Name")).filter(v=>v));
    const seen=new Set();
    FILTERED.filter(r=>isYes(g(r,"Critical Position"))&&g(r,"Successor Name")&&!talentNames.has(g(r,"Successor Name")))
      .forEach(r=>{
        const n=g(r,"Successor Name");
        const rowObj={name:n||"—",role:g(r,"Successor Position"),family:g(r,"Job Family Successor"),
          level:g(r,"Level Successor"),bu:g(r,"Business Unit Successor"),
          readiness:g(r,"Final Score Readiness Text"),flightRisk:g(r,"Successor Flight Risk"),clevel:"",critical:""};
        allRows.push(rowObj);
        if(!n||seen.has(n))return; seen.add(n); rows.push(rowObj);
      });
    rows.sort((a,b)=>a.name.localeCompare(b.name));
    allRows.sort((a,b)=>a.name.localeCompare(b.name));
  } else if(key==="successor"){
    title="Successor";sub="All identified successors";
    const seen=new Set();
    FILTERED.filter(r=>g(r,"Successor Name")).forEach(r=>{
      const n=g(r,"Successor Name");
      const nk=(n||"").toLowerCase().trim();
      const rowObj={name:n||"—",role:g(r,"Position Name")+" (target)",family:g(r,"Job Family Successor"),
        level:g(r,"Level Successor"),bu:g(r,"Business Unit Successor"),
        readiness:g(r,"Final Score Readiness Text"),flightRisk:g(r,"Successor Flight Risk"),
        clevel:isYes(g(r,"C-Level"))?"Yes":"",critical:isYes(g(r,"Critical Position"))?"Yes":""};
      allRows.push(rowObj);
      if(!nk||seen.has(nk))return; seen.add(nk); rows.push(rowObj);
    });
  } else if(key==="cposition"){
    title="C-Position";sub="All C-Level positions";
    const pm2=buildPM(FILTERED);
    Object.values(pm2).filter(p=>p.clevel).forEach(p=>{
      const rowObj={name:p.incumbent||"—",role:p.title||"—",level:p.level,bu:p.bu,clevel:"Yes"};
      rows.push(rowObj); allRows.push(rowObj);
    });
  } else if(key==="cposition_with_succ"){
    title="C-Position with Successor";sub="C-Level positions that have at least one successor";
    const seen=new Set();
    FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")).forEach(r=>{
      const n=g(r,"Successor Name");
      const rowObj={name:n||"—",role:g(r,"Position Name"),level:g(r,"Level Successor"),bu:g(r,"Business Unit Successor"),
        readiness:g(r,"Final Score Readiness Text"),flightRisk:g(r,"Successor Flight Risk"),clevel:"Yes"};
      allRows.push(rowObj);
      if(!n||seen.has(n))return; seen.add(n); rows.push(rowObj);
    });
  } else if(key==="cposition_no_succ"){
    title="C-Position without Successor";sub="C-Level positions with no successor assigned";
    const pm2=buildPM(FILTERED);
    Object.values(pm2).filter(p=>p.clevel&&p.successors.length===0).forEach(p=>{
      const rowObj={name:p.incumbent||"—",role:p.title||"—",level:p.level,bu:p.bu,clevel:"Yes"};
      rows.push(rowObj); allRows.push(rowObj);
    });
  } else if(key==="cplan_single"){
    title="With Single Plan";sub="C-Level positions with exactly one successor";
    // Build map: Position ID → { row, distinct successors, readiness values }
    const _s1PosMap={};
    FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")).forEach(r=>{
      const pid=(g(r,"Position ID")||((g(r,"Position Name")||"")+"___"+(g(r,"Employee Name")||""))||"").toLowerCase().trim();
      if(!pid)return;
      if(!_s1PosMap[pid])_s1PosMap[pid]={row:r,succs:[],reads:[]};
      const sn=g(r,"Successor Name"); const sr=g(r,"Final Score Readiness Text");
      if(sn&&!_s1PosMap[pid].succs.includes(sn))_s1PosMap[pid].succs.push(sn);
      if(sr&&!_s1PosMap[pid].reads.includes(sr))_s1PosMap[pid].reads.push(sr);
    });
    // Only positions with exactly 1 distinct successor
    Object.values(_s1PosMap).filter(({succs})=>succs.length===1).forEach(({row,succs,reads})=>{
      const rowObj={name:g(row,"Employee Name")||"—",role:g(row,"Position Name"),level:g(row,"Level"),bu:g(row,"Business Unit Name"),
        successorNames:succs[0],readiness:reads[0]||"",clevel:"Yes"};
      rows.push(rowObj); allRows.push(rowObj);
    });
  } else if(key==="cplan_multiple"){
    title="With Multiple Plan";sub="C-Level positions with 2+ successors — one row per successor";
    // Build set of Position IDs with 2+ distinct successors
    const _smPosSucc={};
    FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")).forEach(r=>{
      const pid=(g(r,"Position ID")||((g(r,"Position Name")||"")+"___"+(g(r,"Employee Name")||""))||"").toLowerCase().trim();
      const sk=(g(r,"Successor Name")||"").toLowerCase().trim();
      if(!pid||!sk)return;
      if(!_smPosSucc[pid])_smPosSucc[pid]=new Set();
      _smPosSucc[pid].add(sk);
    });
    // One row per unique (position, successor) pair — positions with 2+ successors repeat
    const seenPair=new Set();
    FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")).forEach(r=>{
      const pid=(g(r,"Position ID")||((g(r,"Position Name")||"")+"___"+(g(r,"Employee Name")||""))||"").toLowerCase().trim();
      const sk=(g(r,"Successor Name")||"").toLowerCase().trim();
      if(!pid||!sk||!_smPosSucc[pid]||_smPosSucc[pid].size<2)return;
      const pair=pid+'|||'+sk;
      if(seenPair.has(pair))return; seenPair.add(pair);
      const rowObj={name:g(r,"Employee Name")||"—",role:g(r,"Position Name"),level:g(r,"Level"),bu:g(r,"Business Unit Name"),
        successorNames:g(r,"Successor Name")||"",readiness:g(r,"Final Score Readiness Text"),clevel:"Yes"};
      rows.push(rowObj); allRows.push(rowObj);
    });
  } else if(key==="avail_cl"){
    title="C-Position Coverage";sub="C-Level succession plan rows (150 covered of 163 total)";
    // Only include rows WITH a successor — these are the 150 plan rows
    // _key = Position ID for correct unique/non-unique dedup (matches 57/93 split)
    FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")).forEach(r=>{
      const posId=g(r,"Position ID")||((g(r,"Position Name")||"")+"___"+(g(r,"Employee Name")||""));
      const rowObj={
        _key:posId,
        name:g(r,"Position Name")||"—",
        role:g(r,"Employee Name")||"—",
        level:g(r,"Level"),
        bu:g(r,"Business Unit Name"),
        successorNames:g(r,"Successor Name")||"",
        readiness:g(r,"Final Score Readiness Text")||"",
        clevel:"Yes",
        hasSuccessor:true
      };
      allRows.push(rowObj); rows.push(rowObj);
    });
  } else if(key==="avail_noncl"){
    title="Non C-Position Coverage";sub="Non C-Level succession plan rows (332 covered of 437 total)";
    // Only include rows WITH a successor — these are the 332 plan rows
    FILTERED.filter(r=>!isYes(g(r,"C-Level"))&&g(r,"Successor Name")).forEach(r=>{
      const posId=g(r,"Position ID")||((g(r,"Position Name")||"")+"___"+(g(r,"Employee Name")||""));
      const rowObj={
        _key:posId,
        name:g(r,"Position Name")||"—",
        role:g(r,"Employee Name")||"—",
        level:g(r,"Level"),
        bu:g(r,"Business Unit Name"),
        successorNames:g(r,"Successor Name")||"",
        readiness:g(r,"Final Score Readiness Text")||"",
        clevel:"",
        hasSuccessor:true
      };
      allRows.push(rowObj); rows.push(rowObj);
    });
  }

  SNAP_MODAL_DATA=rows;
  SNAP_MODAL_ALL_DATA=allRows;
  SNAP_FILTER='unique';
  document.getElementById("snapModalTitle").textContent=title;
  document.getElementById("snapModalSub").textContent=sub;
  document.getElementById("snapModalSearch").value="";
  // Single source of truth: let setSnapFilter('unique') render the table.
  // Previously renderSnapModalTable(rows) ran AFTER setSnapFilter and overwrote
  // the 216-row "Unique" view with the 370-row deduped-by-name view, while the
  // dropdown still said "Unique" — causing the count/dropdown mismatch.
  // Reset successor filter each time modal opens
  SNAP_SUCC_FILTER='all';
  document.querySelectorAll('.snap-succ-btn').forEach(b=>{
    const on=b.dataset.f==='all';
    b.style.background=on?'var(--navy)':'var(--skyxs)';
    b.style.color=on?'#fff':'var(--navy)';
    b.style.fontWeight=on?'600':'500';
  });
  const _sfRow=document.getElementById('snapSuccFilterRow');
  if(_sfRow)_sfRow.style.display=(key==='talent'||key==='employee'||key==='avail_cl'||key==='avail_noncl')?'flex':'none';
  // Position-centric views: rows are already per unique Position ID from buildPM,
  // so unique/non-unique by name doesn't apply — always show all rows.
  // C-Level plan/successor views: card totals are all-rows counts, so match with 'all'.
  const _posKeys=['cposition','cposition_with_succ','cposition_no_succ','cplan_single','cplan_multiple'];
  const _isPosView=_posKeys.includes(key);
  const _filterSel=document.getElementById('snapFilterSelect');
  const _filterLabel=_filterSel&&_filterSel.previousElementSibling;
  if(_filterSel){_filterSel.style.display=_isPosView?'none':'';if(_filterLabel&&_filterLabel.tagName==='SPAN')_filterLabel.style.display=_isPosView?'none':'';}
  setSnapFilter(_isPosView?'all':'unique');
  const panel=document.getElementById("snapInlinePanel");
  // highlight active card
  document.querySelectorAll(".snap-card").forEach(c=>c.style.outline="none");
  const activeCard=document.getElementById("snap-card-"+key);
  if(activeCard)activeCard.style.outline="2px solid #60A5FA";
  panel.style.display="block";
  panel.scrollIntoView({behavior:"smooth",block:"start"});
}

const SNAP_PER_PAGE=10;
let SNAP_PAGE=1;
let SNAP_CURRENT_ROWS=[];
let SNAP_MODAL_ALL_DATA=[];
let SNAP_FILTER='unique';
let SNAP_KEY='';
let SNAP_SUCC_FILTER='all';

function setSnapSuccFilter(f){
  SNAP_SUCC_FILTER=f;
  document.querySelectorAll('.snap-succ-btn').forEach(b=>{
    const on=b.dataset.f===f;
    b.style.background=on?'var(--navy)':'var(--skyxs)';
    b.style.color=on?'#fff':'var(--navy)';
    b.style.fontWeight=on?'600':'500';
  });
  setSnapFilter(SNAP_FILTER);
}

function setSnapFilter(f){
  SNAP_FILTER=f;
  const dd=document.getElementById('snapFilterSelect');
  if(dd&&dd.value!==f)dd.value=f;
  // Build dedup-key frequency map — use _key if set (e.g. Position ID), else fall back to name
  const nameCount={};
  SNAP_MODAL_ALL_DATA.forEach(r=>{
    const k=((r._key||r.name)||"").toLowerCase().trim();
    if(k)nameCount[k]=(nameCount[k]||0)+1;
  });
  const _dk=r=>((r._key||r.name)||"").toLowerCase().trim();
  let base;
  if(f==='unique')
    base=SNAP_MODAL_ALL_DATA.filter(r=>nameCount[_dk(r)]===1);
  else if(f==='nonunique')
    base=SNAP_MODAL_ALL_DATA.filter(r=>nameCount[_dk(r)]>1);
  else
    base=SNAP_MODAL_ALL_DATA;
  if(SNAP_SUCC_FILTER==='with')base=base.filter(r=>r.hasSuccessor===true);
  else if(SNAP_SUCC_FILTER==='without')base=base.filter(r=>r.hasSuccessor===false);
  // Re-apply search
  const q=(document.getElementById("snapModalSearch")||{}).value||"";
  const filtered=q?base.filter(r=>Object.values(r).some(v=>(v||"").toString().toLowerCase().includes(q.toLowerCase()))):base;
  SNAP_PAGE=1;
  SNAP_CURRENT_ROWS=filtered;
  _renderSnapPage();
}

function renderSnapModalTable(rows){
  SNAP_CURRENT_ROWS=rows;
  SNAP_PAGE=1;
  _renderSnapPage();
}

function _renderSnapPage(){
  // Sort by name so rows with the same name sit next to each other.
  // Stable sort + localeCompare keeps the visual grouping consistent.
  SNAP_CURRENT_ROWS.sort((a,b)=>(a.name||"").localeCompare(b.name||""));
  const rows=SNAP_CURRENT_ROWS;
  const total=rows.length;
  const totalPages=Math.max(1,Math.ceil(total/SNAP_PER_PAGE));
  if(SNAP_PAGE>totalPages)SNAP_PAGE=totalPages;
  const start=(SNAP_PAGE-1)*SNAP_PER_PAGE;
  const slice=rows.slice(start,start+SNAP_PER_PAGE);

  const hasReadiness=rows.some(r=>r.readiness!==undefined);
  const hasRetRisk=rows.some(r=>r.retRisk!==undefined&&r.retRisk);
  const hasAgeGroup=rows.some(r=>r.ageGroup!==undefined&&r.ageGroup);
  const hasSuccNames=rows.some(r=>r.successorNames&&r.successorNames!=="");
  const hasSuccReadiness=rows.some(r=>r.successorReadiness&&r.successorReadiness!=="");
  const hasCritical=rows.some(r=>r.critical==="Yes");
  const cols=[
    {h:"Name",k:"name",fn:v=>'<span style="font-weight:600;color:var(--navy)">'+v+"</span>"},
    {h:"Role / Target",k:"role"},
    {h:"Job Family",k:"family"},
    {h:"Level",k:"level"},
    {h:"Business Unit",k:"bu"},
    ...(hasSuccNames?[{h:"Successor(s)",k:"successorNames",fn:v=>v?'<span style="color:var(--mid);font-size:11px;">'+v+'</span>':"—"}]:[]),
    ...(hasSuccReadiness?[{h:"Successor Readiness",k:"successorReadiness",fn:v=>v?v.split(", ").map(r=>'<span class="b '+rColor(r)+'" style="font-size:9px;margin-right:2px;">'+r+'</span>').join(""):"—"}]:[]),
    ...(hasReadiness?[{h:"Readiness",k:"readiness",fn:v=>v?'<span class="b '+rColor(v)+'">'+v+"</span>":"—"}]:[]),
    ...(hasRetRisk?[{h:"Retire Risk",k:"retRisk",fn:v=>{const cl=v==="Imminent"||v==="High"?"b-red":v==="Moderate"?"b-amber":"b-green";return v?'<span class="b '+cl+'">'+v+"</span>":"—";}}]:[]),
    ...(hasAgeGroup?[{h:"Age Group",k:"ageGroup"}]:[]),
    {h:"C-Level",k:"clevel",fn:v=>v?'<span class="b b-gold">Yes</span>':"—"},
    ...(hasCritical?[{h:"Critical",k:"critical",fn:v=>v?'<span class="b b-red">Yes</span>':"—"}]:[]),
  ];

  document.getElementById("snapModalCount").textContent=
    "Showing "+(total===0?0:start+1)+"–"+Math.min(start+SNAP_PER_PAGE,total)+" of "+total+" record"+(total!==1?"s":"");

  const head=document.getElementById("snapModalHead");
  const body=document.getElementById("snapModalBody");
  head.innerHTML="<tr>"+cols.map(c=>"<th>"+c.h+"</th>").join("")+"</tr>";
  if(!slice.length){body.innerHTML='<tr><td colspan="'+cols.length+'" class="snap-modal-empty">No records found.</td></tr>';}
  else{body.innerHTML=slice.map(row=>"<tr>"+cols.map(c=>{const v=row[c.k]||"";return "<td>"+(c.fn?c.fn(v):(v||"—"))+"</td>";}).join("")+"</tr>").join("");}

  // pagination bar
  renderPagination("snapPgBar",SNAP_PAGE,totalPages,total,start,Math.min(start+SNAP_PER_PAGE,total),p=>{SNAP_PAGE=p;_renderSnapPage();});
}

function filterSnapModal(){
  setSnapFilter(SNAP_FILTER);
}

function closeSnapModal(){
  const panel=document.getElementById("snapInlinePanel");
  if(panel)panel.style.display="none";
  document.getElementById("snapModalSearch").value="";
  document.querySelectorAll(".snap-card").forEach(c=>c.style.outline="none");
}
document.addEventListener("keydown",e=>{if(e.key==="Escape")closeSnapModal();});


// ── SUCCESSION POSTURE ──
function renderPosture(){
  const pm=buildPM(FILTERED);const arr=Object.values(pm);
  const totalEmp=new Set(FILTERED.map(r=>g(r,"Employee Name")).filter(v=>v)).size;

  // 1. Talent as Successor: unique successor names where Successor As Talent = YES
  const talentSuccSet=new Set(FILTERED.filter(r=>isYes(g(r,"Successor As Talent"))).map(r=>g(r,"Successor Name")).filter(v=>v));
  const talentSucc=talentSuccSet.size;

  // 2. New Potential Talent: unique Successor Name where As Talent Calibrated is NOT "Yes"
  const talentIncSet=new Set(FILTERED.filter(r=>isYes(g(r,"As Talent Calibrated"))).map(r=>g(r,"Employee Name")).filter(v=>v));
  const allSuccSet=new Set(FILTERED.map(r=>g(r,"Successor Name")).filter(v=>v));
  const newPotTalent=new Set(
    FILTERED.filter(r=>g(r,"Successor Name")&&!isYes(g(r,"As Talent Calibrated")))
            .map(r=>g(r,"Successor Name"))
  ).size;

  // 3. Hidden Gem: unique successors to a Critical Position
  const hiddenGemSet=new Set(
    FILTERED.filter(r=>isYes(g(r,"Critical Position"))&&g(r,"Successor Name"))
            .map(r=>g(r,"Successor Name"))
  );
  const hiddenGem=hiddenGemSet.size;

  // 4. New Potential Successor: unique Employee Names who are neither talent nor a hidden gem nor already a successor
  const allEmpSet=new Set(FILTERED.map(r=>g(r,"Employee Name")).filter(v=>v));
  const newPotSucc=new Set([...allEmpSet].filter(n=>!talentIncSet.has(n)&&!hiddenGemSet.has(n)&&!allSuccSet.has(n))).size;

  const pct=(n)=>totalEmp?(n/totalEmp*100).toFixed(1)+"%":"";

  // unknown/dupe for Talent as Successor
  const talentSuccUnknown=FILTERED.filter(r=>isYes(g(r,"Successor As Talent"))&&!g(r,"Successor Name")).length;
  const talentSuccKnown=FILTERED.filter(r=>isYes(g(r,"Successor As Talent"))&&g(r,"Successor Name")).length;
  const talentSuccDupes=talentSuccKnown-talentSucc;

  // unknown/dupe for New Potential Talent — rows where As Talent Calibrated is No/blank and has a successor
  const nptRows=FILTERED.filter(r=>g(r,"Successor Name")&&!isYes(g(r,"As Talent Calibrated")));
  const nptUnknown=FILTERED.filter(r=>!g(r,"Successor Name")&&!isYes(g(r,"As Talent Calibrated"))).length;
  const nptDupes=Math.max(0,nptRows.length-newPotTalent);

  const items=[
    {key:"successor_as_talent",label:"Successor as Talent",sub:"Successors flagged as talent",val:talentSucc,pct:pct(talentSucc),unknown:talentSuccUnknown,dupes:Math.max(0,talentSuccDupes)},
    {key:"new_potential_talent",label:"New Potential Talent",sub:"Successors not yet flagged as talent",val:newPotTalent,pct:pct(newPotTalent),unknown:nptUnknown,dupes:Math.max(0,nptDupes)},
  ];
  document.getElementById("postureItems").innerHTML='<div class="posture-tiles">'+items.map(it=>`
    <div class="posture-tile" style="display:block;padding:10px 12px;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div class="posture-tile-label">${it.label}</div>
          <div class="posture-tile-sub">${it.sub}</div>
        </div>
        <div class="posture-tile-val" style="font-size:26px;">${(it.val+it.dupes+it.unknown).toLocaleString()}</div>
      </div>
      <div style="margin-top:6px;display:flex;flex-direction:column;gap:3px;border-top:1px solid var(--border);padding-top:6px;">
        <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unique</span><span style="font-weight:700;color:var(--navy);">${it.val.toLocaleString()}</span></div>
        <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unknown</span><span style="font-weight:700;color:var(--navy);">${it.unknown.toLocaleString()}</span></div>
        <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non-Unique</span><span style="font-weight:700;color:var(--navy);">${it.dupes.toLocaleString()}</span></div>
      </div>
      <button class="rd-view-btn" onclick="openSnapModal('${it.key}')">View Details →</button>
    </div>`).join("")+'</div>';
}

// ── C-LEVEL POSTURE ──
function renderCLevelPosture(){
  const pm=buildPM(FILTERED);const arr=Object.values(pm);
  const clArr=arr.filter(p=>p.clevel);
  const clPos=clArr.length;
  const clWithSucc=clArr.filter(p=>p.successors.length>0).length;
  const clSingle=clArr.filter(p=>p.successors.length===1).length;
  const clMultiple=clArr.filter(p=>p.successors.length>1).length;

  // Successor to be C-Level: unique vs non-unique
  const clSuccRows=FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const clSuccNames=clSuccRows.map(r=>g(r,"Successor Name")).filter(v=>v);
  const clSuccUnique=new Set(clSuccNames).size;
  const clSuccNonUnique=clSuccNames.length-clSuccUnique;

  // C-Position: unique (distinct position IDs) vs non-unique (duplicate rows)
  const clPosNames=FILTERED.filter(r=>isYes(g(r,"C-Level"))).map(r=>g(r,"Position ID")).filter(v=>v);
  const clPosUnique=new Set(clPosNames).size;
  const clPosNonUnique=clPosNames.length-clPosUnique;

  // C-Position with Successor: unique positions with ≥1 successor vs extra rows
  const clWithSuccRows=FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const clWithSuccPosNames=clWithSuccRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const clWithSuccUnique=new Set(clWithSuccPosNames).size;
  const clWithSuccNonUnique=clWithSuccPosNames.length-clWithSuccUnique;

  // Single/Multiple/Unknown plan counts (positions with 0 successor = no plan, not unknown)
  // Unknown here = C-Level positions where successor name is empty/blank
  const clNoSucc=clArr.filter(p=>p.successors.length===0).length;

  // Single Plan   = C-Level positions with exactly 1 successor
  // Multiple Plan = total successor rows from C-Level positions with 2+ successors
  const clSingleTotal  =clSingle;
  const clMultipleTotal=clArr.filter(p=>p.successors.length>1).reduce((acc,p)=>acc+p.successors.length,0);

  // Sub-breakdown helper matching posture tile style
  const subBreak=(u,n)=>`
    <div style="margin-top:6px;display:flex;flex-direction:column;gap:3px;border-top:1px solid var(--border);padding-top:6px;">
      <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unique</span><span style="font-weight:700;color:var(--navy);">${u.toLocaleString()}</span></div>
      <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non-Unique</span><span style="font-weight:700;color:var(--navy);">${n.toLocaleString()}</span></div>
    </div>`;
  const subBreakWithUnk=(u,n,unk)=>`
    <div style="margin-top:6px;display:flex;flex-direction:column;gap:3px;border-top:1px solid var(--border);padding-top:6px;">
      <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unique</span><span style="font-weight:700;color:var(--navy);">${u.toLocaleString()}</span></div>
      <div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non-Unique</span><span style="font-weight:700;color:var(--navy);">${n.toLocaleString()}</span></div>
      ${unk>0?`<div style="font-size:12px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unknown</span><span style="font-weight:700;color:var(--navy);">${unk.toLocaleString()}</span></div>`:""}
    </div>`;
  const tileWithBreak=(label,total,breakdown,key='')=>`
    <div class="clevel-tile" style="display:block;padding:10px 12px;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div class="clevel-tile-label" style="font-size:12px;font-weight:600;color:var(--navy);">${label}</div>
        <div class="clevel-tile-val" style="font-size:26px;">${total}</div>
      </div>
      ${breakdown}
      ${key?`<button class="rd-view-btn" onclick="openSnapModal('${key}')">View Details →</button>`:''}
    </div>`;

  document.getElementById("clevelByPos").innerHTML=`
    <div class="clevel-tiles">
      <div class="clevel-section-lbl">By Position</div>
      ${tileWithBreak("C-Position",(clPosUnique+clPosNonUnique).toLocaleString(),subBreak(clPosUnique,clPosNonUnique),"cposition")}
      ${tileWithBreak("C-Position with Successor",(clWithSuccUnique+clWithSuccNonUnique).toLocaleString(),subBreak(clWithSuccUnique,clWithSuccNonUnique),"cposition_with_succ")}
      ${tileWithBreak("C-Position without Successor",clNoSucc.toLocaleString(),"","cposition_no_succ")}
      <div class="clevel-section-lbl" style="margin-top:4px">By Employee</div>
      ${tileWithBreak("Successor to be C-Level",(clSuccUnique+clSuccNonUnique).toLocaleString(),subBreak(clSuccUnique,clSuccNonUnique),"successor")}
      ${tileWithBreak("With Single Plan",clSingleTotal.toLocaleString(),"","cplan_single")}
      ${tileWithBreak("With Multiple Plan",clMultipleTotal.toLocaleString(),"","cplan_multiple")}
    </div>`;
  document.getElementById("clevelByEmp").innerHTML="";
}

// ── READINESS POSTURE ──
function renderReadinessPosture(){
  const clRows=FILTERED.filter(r=>isYes(g(r,"C-Level")));
  const nonClRows=FILTERED.filter(r=>!isYes(g(r,"C-Level")));
  const clPosMap=buildPM(clRows);
  const clPosCount=Object.keys(clPosMap).length||1;

  // C-Level readiness
  const clNow=clRows.filter(r=>(g(r,"Final Score Readiness Text")||"").toLowerCase().includes("now")).length;
  const clLater=clRows.filter(r=>{const v=(g(r,"Final Score Readiness Text")||"").toLowerCase();return v.includes("1-3")||v.includes("later");}).length;
  const clFuture=clRows.filter(r=>{const v=(g(r,"Final Score Readiness Text")||"").toLowerCase();return v.includes("3-5")||v.includes("future");}).length;

  // Non-C-Level readiness
  const nonClNow=nonClRows.filter(r=>(g(r,"Final Score Readiness Text")||"").toLowerCase().includes("now")).length;
  const nonClLater=nonClRows.filter(r=>{const v=(g(r,"Final Score Readiness Text")||"").toLowerCase();return v.includes("1-3")||v.includes("later");}).length;
  const nonClFuture=nonClRows.filter(r=>{const v=(g(r,"Final Score Readiness Text")||"").toLowerCase();return v.includes("3-5")||v.includes("future");}).length;

  // All-rows totals
  const readyNow=clNow+nonClNow;
  const readyLater=clLater+nonClLater;
  const readyFuture=clFuture+nonClFuture;

  // Total Plan = C-Position Coverage numerator + Non C-Position Coverage numerator
  // (all rows that have a Successor Name, split by C-Level flag)
  const _clPlanRows   =FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const _nonClPlanRows=FILTERED.filter(r=>!isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const clPlanCount   =_clPlanRows.length;    // matches C-Position Coverage numerator
  const nonClPlanCount=_nonClPlanRows.length; // matches Non C-Position Coverage numerator
  const totalPlan     =clPlanCount+nonClPlanCount;
  const unknownPlanCount=0;
  // Unique/Non-Unique by Position ID — mirrors the Coverage tile breakdown
  const _clPlanPosIds   =_clPlanRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const clPlanUnique    =new Set(_clPlanPosIds).size;
  const clPlanNonUnique =_clPlanPosIds.length-clPlanUnique;
  const _nonClPlanPosIds   =_nonClPlanRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const nonClPlanUnique    =new Set(_nonClPlanPosIds).size;
  const nonClPlanNonUnique =_nonClPlanPosIds.length-nonClPlanUnique;

  const totalForBar=readyNow+readyLater+readyFuture||1;
  const cards=[
    {lbl:"Ready Within a Year",sub:"All successors · Ready Now",val:readyNow,cls:"rn",bench:"",clV:clNow,nonClV:nonClNow},
    {lbl:"Ready Later (1–3 yrs)",sub:"All successors · Ready Later",val:readyLater,cls:"rl",bench:"",clV:clLater,nonClV:nonClLater},
    {lbl:"Ready Future (>3 yrs)",sub:"All successors · Ready Future",val:readyFuture,cls:"rf",bench:"",clV:clFuture,nonClV:nonClFuture},
    {lbl:"Total Plan",sub:"All succession plans",val:totalPlan,cls:"tot",bench:"",
      clPlan:clPlanCount,nonClPlan:nonClPlanCount,unknownPlan:unknownPlanCount,
      clPlanUnique:clPlanUnique,clPlanNonUnique:clPlanNonUnique,
      nonClPlanUnique:nonClPlanUnique,nonClPlanNonUnique:nonClPlanNonUnique},
  ];
  const subRow=(label,val,color)=>`<div style="font-size:10px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;padding-left:12px;"><span>${label}</span><span style="font-weight:600;color:${color||'var(--t2)'}">${val.toLocaleString()}</span></div>`;
  const subDetail=(c)=>{
    if(c.cls==="tot") return `<div style="margin-top:6px;display:flex;flex-direction:column;gap:3px;border-top:1px solid var(--border);padding-top:6px;">
      <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>C-Level Plans</span><span style="font-weight:700;color:var(--mid)">${c.clPlan.toLocaleString()}</span></div>
      ${subRow("Unique",c.clPlanUnique)}
      ${subRow("Non-Unique",c.clPlanNonUnique)}
      <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non C-Level Plans</span><span style="font-weight:700;color:var(--t2)">${c.nonClPlan.toLocaleString()}</span></div>
      ${subRow("Unique",c.nonClPlanUnique)}
      ${subRow("Non-Unique",c.nonClPlanNonUnique)}
    </div>`;
    return `<div style="margin-top:6px;display:flex;flex-direction:column;gap:3px;border-top:1px solid var(--border);padding-top:6px;">
      <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>C-Level</span><span style="font-weight:700;color:var(--mid)">${c.clV.toLocaleString()}</span></div>
      <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non C-Level</span><span style="font-weight:700;color:var(--t2)">${c.nonClV.toLocaleString()}</span></div>
    </div>`;
  };
  document.getElementById("readinessItems").innerHTML='<div class="readiness-grid">'+cards.map(c=>`
    <div class="readiness-card${c.cls==="tot"?" total":""}" style="display:block;padding:10px 12px;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div class="readiness-card-lbl">${c.lbl}</div>
          <div class="readiness-card-sub">${c.sub}</div>
        </div>
        <div class="readiness-card-val ${c.cls}" style="position:static;">${c.val}</div>
      </div>
      ${subDetail(c)}
      ${c.cls!=="tot"?`<button class="rd-view-btn" onclick="openReadinessDetail('${c.cls}')">View Details →</button>`:""}
    </div>`).join("")+'</div>';
}

// ── BENCH STRENGTH ──
function renderBench(){
  const pm=buildPM(FILTERED);const arr=Object.values(pm);
  const clArr=arr.filter(p=>p.clevel);const clPos=clArr.length||1;
  // All C-Level rows with a successor name
  const clPlanRows=FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));

  // ── Unique view: one entry per Position ID, use that position's first row for calibration status
  const posCalibMap={};
  clPlanRows.forEach(r=>{
    const pid=g(r,"Position ID");
    if(pid&&!posCalibMap[pid])posCalibMap[pid]=(g(r,"Calibration Status")||"").toLowerCase().trim();
  });
  const posCalibVals=Object.values(posCalibMap);
  const calibrated=posCalibVals.filter(v=>v==="calibrated").length;
  const nonCalib=posCalibVals.filter(v=>v&&v!=="calibrated").length;
  const unknown=posCalibVals.filter(v=>!v).length;
  const totalPlans=clArr.filter(p=>p.successors.length>0).length;
  const bench=(totalPlans/clPos).toFixed(2);

  // ── Non-Unique view: only the "extra" rows beyond the first per Position ID.
  // Matches the Non-Unique count in the "C-Position with Successor" tile above (e.g. 53 of 150).
  const seenPosIds=new Set();
  const nonUniquePlanRows=clPlanRows.filter(r=>{
    const pid=g(r,"Position ID");
    if(!pid)return false;
    if(seenPosIds.has(pid))return true; // duplicate row -> this is a non-unique entry
    seenPosIds.add(pid);
    return false; // first occurrence -> belongs to Unique bucket, not Non-Unique
  });
  const calibValsNU=nonUniquePlanRows.map(r=>(g(r,"Calibration Status")||"").toLowerCase().trim());
  const calibratedNU=calibValsNU.filter(v=>v==="calibrated").length;
  const nonCalibNU =calibValsNU.filter(v=>v&&v!=="calibrated").length;
  const unknownNU  =calibValsNU.filter(v=>!v).length;
  const totalPlansNU=nonUniquePlanRows.length;
  // Non-Unique denominator: number of Non-Unique C-Position rows (matches "Non-Unique: 53" in C-Position with Successor tile)
  const clPosNU=totalPlansNU;
  const benchNU=clPosNU>0?(totalPlansNU/clPosNU).toFixed(2):"0.00";

  const clWithSucc=clArr.filter(p=>p.successors.length>0).length;
  // C-Position Coverage: match the C-Level Posture tile totals (raw row counts).
  // Numerator   = C-Position with Successor total (150) = rows where C-Level=Yes AND Successor Name exists
  // Denominator = C-Position total (163)                = rows where C-Level=Yes
  const cPosRowsTotal=FILTERED.filter(r=>isYes(g(r,"C-Level"))).length;
  const cPosRowsWithSucc=FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")).length;
  const avail=cPosRowsTotal>0?((cPosRowsWithSucc/cPosRowsTotal)*100).toFixed(2)+"%":"0%";

  // ── Overall view: 150 plans / 163 positions (matches C-Position Coverage 92.02%).
  const totalPlansAll=cPosRowsWithSucc; // 150
  const clPosAll=cPosRowsTotal;          // 163
  const benchAll=clPosAll>0?(totalPlansAll/clPosAll).toFixed(2):"0.00";

  const detailRow=(label,val,ratio)=>`<div style="display:flex;justify-content:space-between;gap:8px;font-size:11px;"><span style="color:var(--t3)">${label}</span><span style="font-weight:700;color:var(--navy)">${val}${ratio!=null?` &nbsp;<span style="color:#7C3AED;font-size:10px">(1: ${ratio})</span>`:""}</span></div>`;

  // Overall detail — calibration rows omitted (covered by Unique + Non-Unique below)
  const detailAll=`<div style="display:flex;flex-direction:column;gap:4px;">
    ${detailRow("Total C-Level Positions",clPosAll)}
    ${detailRow("Total Succession Plans",totalPlansAll)}
  </div>`;

  // Unique detail
  const detail=`<div style="display:flex;flex-direction:column;gap:4px;">
    ${detailRow("Total C-Level Positions",clPos)}
    ${detailRow("Total Succession Plans",totalPlans)}
    ${detailRow("Calibrated",calibrated,totalPlans>0?(calibrated/totalPlans).toFixed(2):"0.00")}
    ${detailRow("Non-Calibrated",nonCalib,totalPlans>0?(nonCalib/totalPlans).toFixed(2):"0.00")}
    ${unknown?detailRow("Unknown Calibration Status",unknown):""}
  </div>`;

  // Non-Unique detail
  const detailNU=`<div style="display:flex;flex-direction:column;gap:4px;">
    ${detailRow("Total C-Level Positions",clPosNU)}
    ${detailRow("Total Succession Plans",totalPlansNU)}
    ${detailRow("Calibrated",calibratedNU,totalPlansNU>0?(calibratedNU/totalPlansNU).toFixed(2):"0.00")}
    ${detailRow("Non-Calibrated",nonCalibNU,totalPlansNU>0?(nonCalibNU/totalPlansNU).toFixed(2):"0.00")}
    ${unknownNU?detailRow("Unknown Calibration Status",unknownNU):""}
  </div>`;

  const el=document.getElementById("benchStrengthVal");
  if(el)el.textContent='1: '+bench;
  const benchLblEl=document.getElementById("benchDetailArea");
  if(benchLblEl)benchLblEl.innerHTML=detail;
  const elAll=document.getElementById("benchStrengthValAll");
  if(elAll)elAll.textContent='1: '+benchAll;
  const benchLblElAll=document.getElementById("benchDetailAreaAll");
  if(benchLblElAll)benchLblElAll.innerHTML=detailAll;
  const elNU=document.getElementById("benchStrengthValNU");
  if(elNU)elNU.textContent='1: '+benchNU;
  const benchLblElNU=document.getElementById("benchDetailAreaNU");
  if(benchLblElNU)benchLblElNU.innerHTML=detailNU;
  document.getElementById("availVal").textContent=avail;
  document.getElementById("availSub").textContent=cPosRowsWithSucc+" of "+cPosRowsTotal+" C-Level positions covered";
  // C-Position Coverage breakdown:
  //   Unique     = (C-Position with Successor Unique 97) / (C-Position Unique 110)
  //   Non-Unique = (C-Position with Successor Non-Unique 53) / (C-Position Non-Unique 53)
  // Numbers are derived from the same row pools used in the C-Level Posture tiles above.
  const cPosAllRows  =FILTERED.filter(r=>isYes(g(r,"C-Level")));
  const cPosWithSuccRows=FILTERED.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const cPosTotalIds        =cPosAllRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const cPosWithSuccIds     =cPosWithSuccRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const cPosTotalUnique     =new Set(cPosTotalIds).size;
  const cPosTotalNonUnique  =cPosTotalIds.length-cPosTotalUnique;
  const cPosSuccUnique      =new Set(cPosWithSuccIds).size;
  const cPosSuccNonUnique   =cPosWithSuccIds.length-cPosSuccUnique;
  // Unique %  = unique plan rows  / total plan rows  (57/150)
  // Non-Unique% = non-unique plan rows / total plan rows (93/150)
  const uniquePct   =cPosRowsWithSucc>0?((cPosSuccUnique/cPosRowsWithSucc)*100).toFixed(2)+"%":"0%";
  const nonUniquePct=cPosRowsWithSucc>0?((cPosSuccNonUnique/cPosRowsWithSucc)*100).toFixed(2)+"%":"0%";
  const clBd=document.getElementById("clAvailBreakdown");
  if(clBd){
    clBd.style.cssText="margin-top:6px;display:flex;flex-direction:column;gap:3px;border-top:1px solid var(--border);padding-top:6px;";
    clBd.innerHTML=`
      <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unique</span><span style="font-weight:700;color:var(--mid)">${uniquePct}</span></div>
      <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non-Unique</span><span style="font-weight:700;color:var(--t2)">${nonUniquePct}</span></div>
    `;
  }
  // Non-C-Level availability — mirrors C-Position Coverage formula:
  //   Main      = (Non-C-Level rows with Successor) / (Non-C-Level rows total)
  //   Unique    = (Unique Non-C Position IDs with Successor) / (Unique Non-C Position IDs total)
  //   Non-Unique= (Non-Unique Non-C Position IDs with Successor) / (Non-Unique Non-C Position IDs total)
  const nonClPosAllRows   =FILTERED.filter(r=>!isYes(g(r,"C-Level")));
  const nonClPosWithSuccRows=nonClPosAllRows.filter(r=>g(r,"Successor Name"));
  const nonClRowsTotal    =nonClPosAllRows.length;
  const nonClRowsWithSucc =nonClPosWithSuccRows.length;
  const nonClAvail=nonClRowsTotal>0?((nonClRowsWithSucc/nonClRowsTotal)*100).toFixed(2)+"%":"0%";
  const nonClEl=document.getElementById("nonClAvailVal");
  const nonClSubEl=document.getElementById("nonClAvailSub");
  if(nonClEl)nonClEl.textContent=nonClAvail;
  if(nonClSubEl)nonClSubEl.textContent=nonClRowsWithSucc+" of "+nonClRowsTotal+" Non C-Level positions covered";
  // Unique / Non-Unique breakdown — same formula style as C-Position Coverage
  const nonClTotalIds       =nonClPosAllRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const nonClSuccIds        =nonClPosWithSuccRows.map(r=>g(r,"Position ID")).filter(v=>v);
  const nonClTotalUnique    =new Set(nonClTotalIds).size;
  const nonClTotalNonUnique =nonClTotalIds.length-nonClTotalUnique;
  const nonClSuccUnique     =new Set(nonClSuccIds).size;
  const nonClSuccNonUnique  =nonClSuccIds.length-nonClSuccUnique;
  // Unique %  = unique plan rows  / total plan rows  (166/332)
  // Non-Unique% = non-unique plan rows / total plan rows (166/332)
  const nonClUniquePct   =nonClRowsWithSucc>0?((nonClSuccUnique/nonClRowsWithSucc)*100).toFixed(2)+"%":"0%";
  const nonClNonUniquePct=nonClRowsWithSucc>0?((nonClSuccNonUnique/nonClRowsWithSucc)*100).toFixed(2)+"%":"0%";
  const nonClBd=document.getElementById("nonClAvailBreakdown");
  if(nonClBd)nonClBd.innerHTML=`
    <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Unique</span><span style="font-weight:700;color:var(--mid)">${nonClUniquePct}</span></div>
    <div style="font-size:11px;color:var(--t3);display:flex;justify-content:space-between;gap:8px;"><span>Non-Unique</span><span style="font-weight:700;color:var(--t2)">${nonClNonUniquePct}</span></div>
  `;
}

// ── CHARTS ──
const PAL=["#1E6FD9","#3B82F6","#60A5FA","#93C5FD","#0891B2","#7C3AED","#10B981","#F59E0B","#EF4444","#6366F1"];
const CSPAL={"Ready Now":"#10B981","1 Year":"#F59E0B","2 Years":"#3B82F6","3+ Years":"#1B4B8A","Unknown":"#94a3b8","Low":"#10B981","Medium":"#F59E0B","High":"#EF4444"};

// Distinct colors for readiness distribution
const READINESS_COLORS={
  "ready now : less than a year":"#10B981",
  "ready later : 1-3 years":"#F59E0B",
  "ready future : 3-5 years":"#3B82F6",
  "unknown":"#94a3b8"
};
function readinessColor(label){
  const l=(label||"").toLowerCase();
  if(l.includes("now"))return"#10B981";
  if(l.includes("1-3")||l.includes("later"))return"#F59E0B";
  if(l.includes("3-5")||l.includes("future"))return"#6366F1";
  if(l.includes("unknown")||l.includes("not"))return"#94a3b8";
  return CSPAL[label]||PAL[0];
}

// Shared datalabels plugin config for bar charts
function barDatalabels(axis="y"){
  return {
    anchor:"end",align:"end",
    color:"#334e68",
    font:{size:9,family:"Poppins",weight:"600"},
    formatter:v=>v>0?v:"",
    padding:{top:2},
    clip:false
  };
}

function mkChart(id,type,labels,vals,colors,opts={}){
  const el=document.getElementById(id);if(!el)return;
  if(CHARTS[id])CHARTS[id].destroy();
  const ds={data:vals,backgroundColor:colors};
  if(type==="bar"){ds.borderRadius=4;ds.borderSkipped=false;}
  if(type==="doughnut"){ds.borderWidth=2;ds.borderColor="#fff";}

  const isHBar=opts.ia==="y";

  CHARTS[id]=new Chart(el,{type,data:{labels,datasets:[ds]},
    options:{responsive:true,maintainAspectRatio:false,
      layout:{padding:{top:type==="bar"?16:0,right:isHBar?30:0}},
      plugins:{
        legend:opts.legend||{display:false},
        datalabels: type==="bar"?{
          anchor: isHBar?"end":"end",
          align: isHBar?"end":"top",
          color:"#334e68",
          font:{size:9,family:"Poppins",weight:"600"},
          formatter:v=>v>0?v:"",
          clamp:true,
          clip:false,
        }:type==="doughnut"?{
          color:"#fff",
          font:{size:10,family:"Poppins",weight:"700"},
          formatter:(v,ctx)=>{
            const total=ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct=total>0?Math.round(v/total*100):0;
            return v>0?(pct+"%"):"";
          },
          anchor:"center",align:"center",
        }:{display:false}
      },
      ...(type==="bar"?{scales:{
        x:{ticks:{color:"#6b93be",font:{size:9,family:"Poppins"},maxRotation:opts.rot||0},grid:{display:false},border:{color:"rgba(30,111,217,.15)"}},
        y:{ticks:{color:"#6b93be",font:{size:9,family:"Poppins"}},grid:{color:"rgba(30,111,217,.07)"},border:{display:false}}
      }}:{}),
      ...(opts.ia?{indexAxis:opts.ia}:{})
    },
    plugins:[ChartDataLabels]
  });
}

function tabBuFilter(selectId, clSelectId){
  let data=FILTERED;
  const bu=getTabBu(selectId);
  if(bu)data=data.filter(r=>(g(r,"Business Unit Name")||"")===bu);
  if(clSelectId){
    const cl=(document.getElementById(clSelectId)||{}).value||"";
    if(cl)data=data.filter(r=>(g(r,"C-Level")||"").toUpperCase()===cl.toUpperCase());
  }
  return data;
}
function renderBUChart(){
  const DATA=getOverviewData();
  const pm=buildPM(DATA);
  const byFam={};
  Object.values(pm).forEach(p=>{
    const f=p.family;
    if(!f)return; // skip positions with no Job Family Successor — no Unknown bucket
    if(!byFam[f])byFam[f]={succ:0,talent:0};
    byFam[f].succ+=p.successors.length;
    byFam[f].talent+=p.clevel?1:0;
  });
  const entries=sorted(byFam,15);
  // (Unknown bar removed — earlier logic to force it rightmost is no longer needed.)
  const labels=entries.map(e=>e[0]);
  const succVals=entries.map(e=>e[1].succ);
  const talentVals=entries.map(e=>e[1].talent);
  const el=document.getElementById("buChart");if(!el)return;
  if(CHARTS.buChart)CHARTS.buChart.destroy();
  CHARTS.buChart=new Chart(el,{
    type:"bar",
    data:{labels,datasets:[
      {label:"Successor",data:succVals,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:false},
      {label:"Talent",data:talentVals,backgroundColor:"#60A5FA",borderRadius:4,borderSkipped:false}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      layout:{padding:{top:18}},
      plugins:{
        legend:{display:true,position:"top",labels:{font:{size:10,family:"Poppins"},boxWidth:10,color:"#5a7fa8"}},
        datalabels:{
          anchor:"end",align:"top",
          color:"#334e68",
          font:{size:9,family:"Poppins",weight:"600"},
          formatter:v=>v>0?v:"",
          clip:false,
        }
      },
      scales:{
        x:{ticks:{color:"#6b93be",font:{size:9,family:"Poppins"},maxRotation:35},grid:{display:false},border:{color:"rgba(30,111,217,.15)"}},
        y:{ticks:{color:"#6b93be",font:{size:9}},grid:{color:"rgba(30,111,217,.07)"},border:{display:false}}
      }
    },
    plugins:[ChartDataLabels]
  });
}

function renderAgeChart(){

  const TAB_DATA=getOverviewData();  const clSuccRows=TAB_DATA.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const bands={"< 30":0,"30-35":0,"36-40":0,"41-45":0,"46-50":0,"51-55":0,"> 55":0};
  let ageUnknown=0;
  clSuccRows.forEach(r=>{
    const grpAge=(g(r,"Group Age")||"").toString().trim();
    if(!grpAge){ageUnknown++;return;}
    if(grpAge==="< 30"||grpAge==="<30")bands["< 30"]++;
    else if(grpAge==="30-35")bands["30-35"]++;
    else if(grpAge==="36-40")bands["36-40"]++;
    else if(grpAge==="41-45")bands["41-45"]++;
    else if(grpAge==="46-50")bands["46-50"]++;
    else if(grpAge==="51-55")bands["51-55"]++;
    else if(grpAge==="> 50"||grpAge===">50"||grpAge==="> 55"||grpAge===">55")bands["> 55"]++;
    else ageUnknown++;
  });
  const ageLabels=Object.keys(bands);
  const ageVals=Object.values(bands);
  const ageColors=["#93C5FD","#3B82F6","#60A5FA","#0891B2","#1E6FD9","#1B4B8A","#172554"];
  if(ageUnknown>0){ageLabels.push("Unknown");ageVals.push(ageUnknown);ageColors.push("#94A3B8");}
  const ageEl=document.getElementById("ageChart");if(!ageEl)return;
  if(CHARTS["ageChart"])CHARTS["ageChart"].destroy();
  const ageMax=Math.max(...ageVals,1);
  CHARTS["ageChart"]=new Chart(ageEl,{
    type:"bar",
    data:{labels:ageLabels,datasets:[{data:ageVals,backgroundColor:ageColors,borderRadius:4,borderSkipped:false}]},
    options:{
      responsive:true,maintainAspectRatio:false,
      layout:{padding:{top:24}},
      plugins:{
        legend:{display:false},
        datalabels:{anchor:"end",align:"top",color:"#334e68",font:{size:9,family:"Poppins",weight:"600"},formatter:v=>v>0?v:"",clip:false}
      },
      scales:{
        x:{ticks:{color:"#6b93be",font:{size:9,family:"Poppins"},maxRotation:45,minRotation:30,autoSkip:false},grid:{display:false},border:{color:"rgba(30,111,217,.15)"}},
        y:{display:false,max:ageMax*1.45}
      }
    },
    plugins:[ChartDataLabels]
  });
}

function renderLayerChart(){

  const TAB_DATA=getOverviewData();  const m=countBy(TAB_DATA.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name")),"Level Successor");
  delete m["Unknown"];delete m[""];
  const e=sorted(m,8),l=e.map(x=>x[0]),v=e.map(x=>x[1]);
  mkChart("layerChart","doughnut",l,v,PAL.slice(0,l.length),{legend:{position:"right",labels:{font:{size:9,family:"Poppins"},boxWidth:8,padding:6,color:"#5a7fa8"}}});
}

function renderJobFamilyChart(){

  const TAB_DATA=getOverviewData();  const clSuccRows=TAB_DATA.filter(r=>isYes(g(r,"C-Level"))&&g(r,"Successor Name"));
  const m=countBy(clSuccRows,"Job Family Successor");delete m["Unknown"];delete m[""];
  const entries=sorted(m,20);
  const labels=entries.map(e=>e[0]),vals=entries.map(e=>e[1]);
  const wrap=document.getElementById("jobFamilyWrap");
  if(wrap)wrap.style.height=Math.max(200,entries.length*22+60)+"px";
  const el=document.getElementById("jobFamilyChart");if(!el)return;
  if(CHARTS.jobFamilyChart)CHARTS.jobFamilyChart.destroy();
  CHARTS.jobFamilyChart=new Chart(el,{
    type:"bar",data:{labels,datasets:[{data:vals,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:false}]},
    options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,
      layout:{padding:{right:32}},
      plugins:{
        legend:{display:false},
        datalabels:{
          anchor:"end",align:"end",
          color:"#334e68",
          font:{size:9,family:"Poppins",weight:"600"},
          formatter:v=>v>0?v:"",
          clip:false,
        }
      },
      scales:{
        x:{ticks:{color:"#6b93be",font:{size:9,family:"Poppins"}},grid:{color:"rgba(30,111,217,.07)"},border:{display:false}},
        y:{ticks:{color:"#334e68",font:{size:9,family:"Poppins"}},grid:{display:false},border:{color:"rgba(30,111,217,.15)"}}
      }
    },
    plugins:[ChartDataLabels]
  });
}

function renderReadinessChart(){

  const TAB_DATA=getOverviewData();  const m=countBy(TAB_DATA,"Final Score Readiness Text");
  const unknown=(m["Unknown"]||0)+(m[""]||0);
  delete m["Unknown"];delete m[""];
  if(unknown>0)m["Unknown"]=unknown;
  const e=sorted(m),l=e.map(x=>x[0]),v=e.map(x=>x[1]);
  const colors=l.map(x=>x==="Unknown"?"#CBD5E1":readinessColor(x));
  const el=document.getElementById("readinessChart");if(!el)return;
  if(CHARTS.readinessChart)CHARTS.readinessChart.destroy();
  CHARTS.readinessChart=new Chart(el,{
    type:"doughnut",
    data:{labels:l,datasets:[{data:v,backgroundColor:colors,borderWidth:2,borderColor:"#fff"}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:true,position:"bottom",labels:{font:{size:9,family:"Poppins"},boxWidth:8,padding:6,color:"#5a7fa8"}},
        datalabels:{
          color:"#fff",
          font:{size:10,family:"Poppins",weight:"700"},
          formatter:(val,ctx)=>{
            const total=ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct=total>0?Math.round(val/total*100):0;
            return val>0?(pct+"%"):"";
          },
          anchor:"center",align:"center",
        }
      }
    },
    plugins:[ChartDataLabels]
  });
}

function renderFlightChart(){

  const TAB_DATA=getOverviewData();  const m=countBy(TAB_DATA,"As Talent Calibrated");
  const unknown=(m["Unknown"]||0)+(m[""]||0);
  delete m["Unknown"];delete m[""];
  if(unknown>0)m["Unknown"]=unknown;
  const e=sorted(m),l=e.map(x=>x[0]),v=e.map(x=>x[1]);
  const colors=l.map(x=>x.toUpperCase()==="YES"?"#10B981":x==="Unknown"?"#CBD5E1":"#94A3B8");
  mkChart("flightChart","doughnut",l,v,colors,{legend:{position:"bottom",labels:{font:{size:9,family:"Poppins"},boxWidth:8,padding:6,color:"#5a7fa8"}}});
}
function renderCritChart(){

  const TAB_DATA=getOverviewData();  const yes=TAB_DATA.filter(r=>isYes(g(r,"Critical Position"))).length;
  const no=TAB_DATA.filter(r=>{const v=(g(r,"Critical Position")||"").toString().trim().toUpperCase();return v==="NO";}).length;
  const unknown=TAB_DATA.filter(r=>!g(r,"Critical Position")||g(r,"Critical Position").toString().trim()==="").length;
  const labels=[],vals=[],colors=[];
  if(yes>0){labels.push("Critical Position");vals.push(yes);colors.push("#EF4444");}
  if(no>0){labels.push("Non-Critical Position");vals.push(no);colors.push("#10B981");}
  if(unknown>0){labels.push("Unknown");vals.push(unknown);colors.push("#CBD5E1");}
  mkChart("critChart","doughnut",labels,vals,colors,{legend:{position:"bottom",labels:{font:{size:9,family:"Poppins"},boxWidth:8,padding:6,color:"#5a7fa8"}}});
}
function renderRetChart(){

  const TAB_DATA=getOverviewData();
  const cats={"Low":0,"Moderate":0,"High":0,"Unknown":0};
  const _retSeen=new Set();
  TAB_DATA.forEach(r=>{
    const key=g(r,"Position ID")||g(r,"NIK")||g(r,"Employee Name");
    if(!key||_retSeen.has(key))return;
    _retSeen.add(key);
    const rr=retRiskFromGroup(g(r,"Group Age"));
    if(cats[rr]!==undefined)cats[rr]++;else cats["Unknown"]++;
  });
  if(cats["Unknown"]===0)delete cats["Unknown"];
  const retTooltipMap={
    "Low":"Age groups: < 30, 30–35, 36–40",
    "Moderate":"Age groups: 41–45, 46–50",
    "High":"Age groups: 51–55, > 55"
  };
  const retCtx=document.getElementById("retChart");
  if(!retCtx)return;
  if(retCtx._chart)retCtx._chart.destroy();
  retCtx._chart=new Chart(retCtx,{
    type:"bar",
    data:{
      labels:Object.keys(cats),
      datasets:[{
        data:Object.values(cats),
        backgroundColor:Object.keys(cats).map(k=>k==="Low"?"#10B981":k==="Moderate"?"#F59E0B":k==="High"?"#EF4444":"#94A3B8"),
        borderRadius:4,borderSkipped:false
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{
          callbacks:{
            title:ctx=>ctx[0].label,
            afterLabel:ctx=>retTooltipMap[ctx.label]||""
          }
        },
        datalabels:{anchor:"end",align:"end",color:"#1B4B8A",font:{size:11,weight:"700",family:"Poppins"}}
      },
      scales:{x:{grid:{display:false},ticks:{color:"#5a7fa8",font:{size:10,family:"Poppins"}}},
              y:{display:false,grid:{display:false},max:Math.max(...Object.values(cats))*1.3}}
    },
    plugins:[ChartDataLabels]
  });
}

// ── POSITIONS ──
let PF_STATE2="all";
function setPF(btn,f){
  PF_STATE2=f;POS_PAGE=1;document.querySelectorAll(".pf").forEach(b=>b.classList.remove("on"));btn.classList.add("on");
  renderPositions();
}
function updateFolderCounts(){
  const pm=buildPM(FILTERED);const arr=Object.values(pm);
  const setct=(id,n)=>{const el=document.getElementById(id);if(el)el.textContent=n;};
  setct("pfct-all",arr.length);
  setct("pfct-cl",arr.filter(p=>p.clevel).length);
  setct("pfct-crit",arr.filter(p=>p.critical).length);
  setct("pfct-ns",arr.filter(p=>p.successors.length===0).length);
  setct("pfct-risk",arr.filter(p=>(p.flightRisk||"").toLowerCase()==="high").length);
}
function riskScore(p){
  let s=0;
  if(p.successors.length===0)s+=40;
  else if(p.successors.length===1)s+=15;
  if(p.retRisk==="Imminent")s+=30;else if(p.retRisk==="High")s+=20;else if(p.retRisk==="Moderate")s+=10;
  const fr=p.flightRisk.toLowerCase();
  if(fr==="high")s+=20;else if(fr==="medium")s+=10;
  if(p.critical)s+=10;
  if(p.clevel)s+=5;
  return Math.min(s,100);
}
function riskLabel(score){
  if(score>=60)return{cls:"risk-critical",dot:"risk-dot-r",txt:"Critical"};
  if(score>=30)return{cls:"risk-watch",dot:"risk-dot-y",txt:"Watch"};
  return{cls:"risk-healthy",dot:"risk-dot-g",txt:"Healthy"};
}
function benchDepth(p){
  const rd=p.successors.map(s=>(s.readiness||"").toLowerCase());
  return{
    now:rd.filter(r=>r.includes("now")).length,
    y1:rd.filter(r=>r.includes("1-3")||r.includes("later")).length,
    y2:0,
    y3p:rd.filter(r=>r.includes("3-5")||r.includes("future")).length
  };
}
function bestReadiness(p){
  const order=["now","1-3","later","3-5","future"];
  for(const o of order){if(p.successors.some(s=>(s.readiness||"").toLowerCase().includes(o)))return o;}
  return p.successors.length?"unknown":"none";
}
function timeGapHtml(best){
  if(best==="none")return'';
  if(best.includes("now"))return'<span class="time-gap tg-now">Ready Now</span>';
  if(best.includes("1-3")||best.includes("later"))return'<span class="time-gap tg-1y">1–3 Year Gap</span>';
  if(best.includes("3-5")||best.includes("future"))return'<span class="time-gap tg-3p">3–5 Year Gap</span>';
  return'<span class="time-gap tg-3p">3+ Year Gap</span>';
}
function renderPositions(){
  POS_PAGE=1; // reset to page 1 on filter/search change
  updateFolderCounts();
  renderPositionsPage();
  window.GAP_PAGE=1;
  window.HR_PAGE=1;
  window._GAP_FILTERS={};
  window._HR_FILTERS={};
  renderGapHR();
}

function renderPositionsPage(){
  const srcPos=tabBuFilter("posBuFilter");
  const pm=buildPM(srcPos);let arr=Object.values(pm);
  if(PF_STATE2==="cl")arr=arr.filter(p=>p.clevel);
  else if(PF_STATE2==="crit")arr=arr.filter(p=>p.critical);
  else if(PF_STATE2==="ns")arr=arr.filter(p=>p.successors.length===0);
  else if(PF_STATE2==="risk")arr=arr.filter(p=>p.flightRisk.toLowerCase()==="high");
  arr.sort((a,b)=>{if(a.clevel&&!b.clevel)return -1;if(!a.clevel&&b.clevel)return 1;if(a.critical&&!b.critical)return -1;if(!a.critical&&b.critical)return 1;return riskScore(b)-riskScore(a);});

  // apply search
  const qs=(document.getElementById("posSearch")||{}).value||"";
  const qsl=qs.toLowerCase();
  if(qsl) arr=arr.filter(p=>(p.title||"").toLowerCase().includes(qsl)||(p.incumbent||"").toLowerCase().includes(qsl)||(p.family||"").toLowerCase().includes(qsl)||(p.bu||"").toLowerCase().includes(qsl));

  POS_ALL_ARR=arr;
  const total=arr.length;
  const totalPages=Math.max(1,Math.ceil(total/POS_PER_PAGE));
  if(POS_PAGE>totalPages)POS_PAGE=totalPages;
  const start=(POS_PAGE-1)*POS_PER_PAGE;
  const slice=arr.slice(start,start+POS_PER_PAGE);

  const sc=document.getElementById("posSearchCount");
  if(sc)sc.textContent=qsl?total+" result"+(total!==1?"s":"")+" found":total+" position"+(total!==1?"s":"");

  document.getElementById("posGrid").innerHTML='<div class="pos-grid">'+slice.map(p=>{
    const tags=[];
    if(p.clevel)tags.push('<span class="b b-gold">C-Level</span>');
    if(p.critical)tags.push('<span class="b b-red">Critical</span>');
    const age=p.incAge||0;
    const rl=age>=50?{cls:"risk-watch",dot:"risk-dot-y",txt:"Watch"}:{cls:"risk-healthy",dot:"risk-dot-g",txt:"Healthy"};
    const bd=benchDepth(p);
    const best=bestReadiness(p);
    const bdHtml='<div class="bench-depth">'
      +(bd.now?'<span class="bench-tier bt-now">&#10003; Ready Now: '+bd.now+"</span>":"")
      +(bd.y1?'<span class="bench-tier bt-1y">1-3 yrs: '+bd.y1+"</span>":"")
      +(bd.y3p?'<span class="bench-tier bt-3p">3-5 yrs: '+bd.y3p+"</span>":"")

      +"</div>";
    const rows=p.successors.length
      ?p.successors.slice(0,4).map(s=>{
        const load=(SUCC_LOAD_MAP[s.name]||[]).length;
        const countBubble=load>1?'<span style="font-size:9px;font-weight:700;background:rgba(30,111,217,.12);color:var(--mid);border-radius:10px;padding:1px 6px;margin-left:5px;flex-shrink:0;">'+load+' pos</span>':"";
        return '<div class="srow">'
          +'<span style="display:flex;align-items:center;gap:2px;">'
          +'<span class="srow-name">'+s.name+'</span>'
          +countBubble
          +'</span>'
          +'<div class="srow-r">'
          +(s.asTalent?'<span class="b b-green" style="font-size:9px">Talent</span>':"")
          +(s.readiness?'<span class="b '+rColor(s.readiness)+'" style="font-size:9px">'+s.readiness+"</span>":"")
          +"</div></div>";
      }).join("")+(p.successors.length>4?'<div style="font-size:10px;color:#94a3b8;padding:2px 8px">+'+( p.successors.length-4)+" more</div>":"")
      :'<div class="nosucc">No successor identified</div>';
    const cls=["pos-card",p.clevel?"pcl":p.critical?"pcrit":"",!p.successors.length?"pns":""].filter(Boolean).join(" ");
    return '<div class="'+cls+'" onclick="openPosDetail(\''+p.id+'\')">'
      +'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:6px">'
      +'<div class="pos-name" title="'+p.title+'">'+p.title+'</div>'
      +'</div>'
      +(p.incumbent?'<div style="display:flex;align-items:center;gap:5px;margin-bottom:5px">'
        +'<svg viewBox="0 0 24 24" fill="none" stroke="#1B4B8A" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="width:12px;height:12px;flex-shrink:0"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>'
        +'<span style="font-size:11.5px;font-weight:700;color:#1B4B8A;letter-spacing:.01em"><span style="color:#3B82F6">Incumbent:</span> '+p.incumbent+'</span>'
        +'</div>'
        :'')
      +'<div class="pos-meta">'+[p.family,p.level,p.years?p.years+" yrs":""].filter(Boolean).join(" · ")+'</div>'
      +'<div class="pos-tags" style="display:flex;flex-wrap:wrap;gap:5px;align-items:center;margin-bottom:10px;">'+tags.join("")+timeGapHtml(best)+'</div>'
      +'<div class="pos-sh" style="margin-top:6px"><span class="pos-sh-lbl">Successors</span><span class="pos-sh-ct">'+p.successors.length+" candidate"+(p.successors.length!==1?"s":"")+"</span></div>"
      +bdHtml
      +rows
      +"</div>";
  }).join("")+'</div>'||'<div style="color:#94a3b8;font-style:italic;padding:20px;text-align:center;">No positions match your search.</div>';

  renderPagination("posPgBar", POS_PAGE, totalPages, total, start, Math.min(start+POS_PER_PAGE,total), (p)=>{POS_PAGE=p;renderPositionsPage();});
}

// ── GAP + HIGH-RISK TABLES (Position Coverage tab, respects posBuFilter) ──
if(!window._GAP_FILTERS)window._GAP_FILTERS={};
if(!window._HR_FILTERS) window._HR_FILTERS={};

function _ghFilter(table,key,value){
  if(table==='gap'){window._GAP_FILTERS[key]=value;window.GAP_PAGE=1;}
  else           {window._HR_FILTERS[key]=value; window.HR_PAGE=1;}
  renderGapHR();
}

function _buildGHThead(theadId, cols, allData, filterState){
  const row='<tr>'+cols.map(col=>{
    if(!col.filterable) return `<th>${col.label}</th>`;
    const vals=[...new Set(allData.map(p=>col.val(p)).filter(v=>v&&v!=='—'))].sort();
    const cur=filterState[col.key]||'';
    const opts='<option value="">All</option>'+vals.map(v=>`<option value="${_esc(v)}"${cur===v?' selected':''}>${_esc(v)}</option>`).join('');
    const active=cur?'filtered':'';
    return `<th><div class="th-col"><span>${col.label}</span><select class="th-fsel ${active}" onchange="_ghFilter('${col.table}','${col.key}',this.value)">${opts}</select></div></th>`;
  }).join('')+'</tr>';
  document.getElementById(theadId).innerHTML=row;
}

function _applyGHFilters(arr, filterState, cols){
  return arr.filter(p=>cols.every(col=>{
    if(!col.filterable)return true;
    const f=filterState[col.key];
    if(!f)return true;
    return col.val(p)===f;
  }));
}

function renderGapHR(){
  const srcPos=tabBuFilter("posBuFilter");
  const pm=buildPM(srcPos);
  const GAP_PER=6;
  const HR_PER=6;
  function applyUniqFilter(arr,filterId){
    const mode=(document.getElementById(filterId)||{}).value||"unique";
    if(mode==="all")return arr;
    const seen=new Set();
    const uniq=arr.filter(p=>{if(seen.has(p.title))return false;seen.add(p.title);return true;});
    if(mode==="unique")return uniq;
    const counts={};arr.forEach(p=>{counts[p.title]=(counts[p.title]||0)+1;});
    return arr.filter(p=>counts[p.title]>1);
  }
  const _fr=p=>{const f=(p.flightRisk||'Unknown');return f.charAt(0).toUpperCase()+f.slice(1).toLowerCase();}

  // ── GAPS table ──
  const gapCols=[
    {key:'position',  label:'Position',  filterable:false, table:'gap', val:p=>p.title},
    {key:'incumbent', label:'Incumbent', filterable:false, table:'gap', val:p=>p.incumbent||'—'},
    {key:'level',     label:'Level',     filterable:true,  table:'gap', val:p=>p.level||'—'},
    {key:'clevel',    label:'C-Level',   filterable:true,  table:'gap', val:p=>p.clevel?'Yes':'—'},
    {key:'critical',  label:'Critical',  filterable:true,  table:'gap', val:p=>p.critical?'Yes':'—'},
    {key:'risk',      label:'Risk',      filterable:true,  table:'gap', val:_fr},
  ];
  const allGaps=Object.values(pm).filter(p=>p.successors.length===0);
  _buildGHThead('gapThead',gapCols,allGaps,window._GAP_FILTERS);
  const gaps=_applyGHFilters(applyUniqFilter(allGaps,"gapUniqFilter"),window._GAP_FILTERS,gapCols);
  if(!window.GAP_PAGE)window.GAP_PAGE=1;
  const gTotalPages=Math.max(1,Math.ceil(gaps.length/GAP_PER));
  if(window.GAP_PAGE>gTotalPages)window.GAP_PAGE=1;
  const gStart=(window.GAP_PAGE-1)*GAP_PER;
  const gSlice=gaps.slice(gStart,gStart+GAP_PER);
  document.getElementById("gapBody").innerHTML=gSlice.map(p=>'<tr><td style="font-weight:600;color:var(--navy)">'+p.title+"</td><td style=\"color:var(--t2)\">"+(p.incumbent||"—")+"</td><td>"+(p.level||"—")+"</td><td>"+(p.clevel?'<span class="b b-gold">Yes</span>':"—")+"</td><td>"+(p.critical?'<span class="b b-red">Yes</span>':"—")+"</td><td><span class=\"b "+(_fr(p).toLowerCase()==="high"?"b-red":_fr(p).toLowerCase()==="medium"?"b-amber":"b-green")+'">'+_fr(p)+"</span></td></tr>").join("")||'<tr><td colspan="6" style="color:#94a3b8;font-style:italic;padding:10px">No positions match the current filter.</td></tr>';
  renderPagination("pgGapBar",window.GAP_PAGE,gTotalPages,gaps.length,gStart,Math.min(gStart+GAP_PER,gaps.length),p=>{window.GAP_PAGE=p;renderGapHR();});

  // ── HIGH RISK table ──
  const hrCols=[
    {key:'position',  label:'Position',  filterable:false, table:'hr', val:p=>p.title},
    {key:'level',     label:'Level',     filterable:true,  table:'hr', val:p=>p.level||'—'},
    {key:'clevel',    label:'C-Level',   filterable:true,  table:'hr', val:p=>p.clevel?'Yes':'—'},
    {key:'incumbent', label:'Incumbent', filterable:false, table:'hr', val:p=>p.incumbent||'—'},
    {key:'risk',      label:'Risk',      filterable:true,  table:'hr', val:_fr},
  ];
  const allHR=Object.values(pm).filter(p=>p.critical&&p.successors.length===0);
  _buildGHThead('hrThead',hrCols,allHR,window._HR_FILTERS);
  const hr=_applyGHFilters(applyUniqFilter(allHR,"hrUniqFilter"),window._HR_FILTERS,hrCols);
  if(!window.HR_PAGE)window.HR_PAGE=1;
  const hTotalPages=Math.max(1,Math.ceil(hr.length/HR_PER));
  if(window.HR_PAGE>hTotalPages)window.HR_PAGE=1;
  const hStart=(window.HR_PAGE-1)*HR_PER;
  const hSlice=hr.slice(hStart,hStart+HR_PER);
  document.getElementById("hrBody").innerHTML=hSlice.map(p=>'<tr><td style="font-weight:600;color:var(--red)">'+p.title+"</td><td>"+(p.level||"—")+"</td><td>"+(p.clevel?'<span class="b b-gold">Yes</span>':"—")+"</td><td>"+(p.incumbent||"—")+"</td><td><span class=\"b "+(_fr(p).toLowerCase()==="high"?"b-red":_fr(p).toLowerCase()==="medium"?"b-amber":"b-green")+'">'+_fr(p)+"</span></td></tr>").join("")||'<tr><td colspan="5" style="color:#94a3b8;font-style:italic;padding:10px">No positions match the current filter.</td></tr>';
  renderPagination("pgHRBar",window.HR_PAGE,hTotalPages,hr.length,hStart,Math.min(hStart+HR_PER,hr.length),p=>{window.HR_PAGE=p;renderGapHR();});
}

// ── SIMULATION ──
const RR={"ready now : less than a year":0,"ready later : 1-3 years":1,"ready future : 3-5 years":2,"unknown":99};
const FR={"low":0,"medium":1,"high":2,"unknown":99};
function renderSimPanel(){
  const pm=buildPM(FILTERED);
  const buSel=document.getElementById("simBuFilter");
  const prevBu=buSel?buSel.value:"";
  const buSet=new Set(Object.values(pm).map(p=>p.bu).filter(v=>v));
  if(buSel){
    buSel.innerHTML='<option value="">— All Business Units —</option>';
    [...buSet].sort().forEach(bu=>{const o=document.createElement("option");o.value=bu;o.textContent=bu;buSel.appendChild(o);});
    if(prevBu&&buSet.has(prevBu))buSel.value=prevBu;
  }
  _filterSimPositions(pm);
  scPopulate(pm);
}
function _filterSimPositions(pm){
  if(!pm)pm=buildPM(FILTERED);
  const buSel=document.getElementById("simBuFilter");
  const sel=document.getElementById("simPosSelect");
  const selectedBu=buSel?buSel.value:"";
  const prev=sel.value;
  sel.innerHTML='<option value="">— Select a position —</option>';
  let positions=Object.values(pm);
  if(selectedBu)positions=positions.filter(p=>p.bu===selectedBu);
  positions.sort((a,b)=>a.title.localeCompare(b.title)).forEach(p=>{
    const o=document.createElement("option");o.value=p.id;
    o.textContent=p.title+(p.clevel?" [C-Level]":p.critical?" [Critical]":"")+(p.incumbent?" — "+p.incumbent:"");
    sel.appendChild(o);
  });
  if(prev)sel.value=prev;
}
let _simSuccData=[];
function runSimulation(){
  const pid=document.getElementById("simPosSelect").value;
  const pm=buildPM(FILTERED);const pos=pm[pid];
  if(!pos){document.getElementById("simResult").innerHTML="Please select a position.";return;}
  const el=document.getElementById("simResult");
  if(!pos.successors.length){el.innerHTML='<span class="sim-vacant">⚠ CRITICAL GAP</span> — <strong>'+pos.title+"</strong> has no identified successors.";return;}
  const scored=pos.successors.map(s=>({...s,score:(RR[(s.readiness||"").toLowerCase()]??99)*10+(FR[("").toLowerCase()]??0)})).sort((a,b)=>a.score-b.score);
  const best=scored[0];const others=scored.slice(1);
  _simSuccData=scored;
  el.innerHTML='<p style="margin-bottom:.6rem;font-size:12px;color:#334e68"><strong style="color:var(--navy)">'+pos.title+'</strong> vacancy simulated:</p>'
    +'<div style="background:linear-gradient(135deg,rgba(59,130,246,.08),rgba(96,165,250,.05));border:1px solid rgba(59,130,246,.2);border-radius:var(--r);padding:.8rem 1rem;margin-bottom:3px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px">'
    +'<div><div style="font-size:9px;font-weight:700;color:var(--bmed);letter-spacing:.1em;text-transform:uppercase;margin-bottom:2px">★ Recommended Successor</div>'
    +'<strong style="font-size:14px;color:var(--navy)">'+_esc(best.name)+"</strong></div>"
    +'<div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;">'
    +'<span class="b '+rColor(best.readiness)+'" style="font-size:11px">'+(best.readiness||"Unknown")+'</span>'
    +'<button onclick="_toggleSimDetail(\'sim-d-0\')" class="rd-view-btn" style="margin:0;width:auto;padding:4px 12px;font-size:10px;">See Detail</button>'
    +'</div></div>'
    +'<div id="sim-d-0" style="display:none;margin-bottom:4px;"></div>'
    +(others.length?'<p style="font-size:10px;font-weight:600;color:#6b93be;text-transform:uppercase;letter-spacing:.08em;margin:.5rem 0 .3rem">Other Candidates</p>':"")
    +others.map((s,i)=>
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:5px 9px;border-radius:5px;background:var(--skyxs);border:1px solid var(--border);margin-bottom:3px">'
      +'<span style="font-size:11px;font-weight:500;color:var(--navy)">'+_esc(s.name)+'</span>'
      +'<div style="display:flex;align-items:center;gap:6px;">'
      +'<span class="b '+rColor(s.readiness)+'" style="font-size:9px">'+(s.readiness||"Unknown")+'</span>'
      +'<button onclick="_toggleSimDetail(\'sim-d-'+(i+1)+'\')" style="padding:2px 9px;font-size:9px;font-family:inherit;font-weight:600;background:var(--skyxs);border:1px solid var(--border2);color:var(--mid);border-radius:999px;cursor:pointer;">Detail</button>'
      +'</div></div>'
      +'<div id="sim-d-'+(i+1)+'" style="display:none;margin-bottom:3px;"></div>'
    ).join("");
}
function _toggleSimDetail(panelId){
  const panel=document.getElementById(panelId);
  if(!panel)return;
  if(panel.style.display!=='none'){panel.style.display='none';return;}
  const idx=parseInt(panelId.replace('sim-d-',''))||0;
  const succ=_simSuccData[idx];
  if(!succ)return;
  const learning=_learningForEmployee(succ.succEmp||null,succ.name);
  const statusCls=learning.status==="Completed"?"b-green":learning.status==="In Progress"?"b-amber":"b-gray";
  const courseHtml=learning.courses.length
    ?'<div style="display:flex;flex-direction:column;gap:3px;">'
      +learning.courses.map(c=>'<div style="display:flex;justify-content:space-between;gap:8px;font-size:11px;padding:3px 0;border-bottom:1px solid var(--border);">'
        +'<span style="color:var(--text)">'+_esc(c.name)+'</span>'
        +'<span style="color:var(--t2);flex-shrink:0;font-weight:600;">'+(c.pct!==null?c.pct+'%':'—')+'</span>'
        +'</div>').join("")
      +'</div>'
    :'<span style="font-size:11px;color:var(--t3);">No learning data — readiness based on calibration score only.</span>';
  const facts=[];
  if(succ.readinessScore)facts.push({l:'Readiness Score',v:succ.readinessScore});
  if(succ.jobComplexity)facts.push({l:'Job Complexity',v:succ.jobComplexity});
  if(succ.yearsExp)facts.push({l:'Years of Experience',v:succ.yearsExp});
  if(succ.techKnowledge)facts.push({l:'Technical Knowledge',v:succ.techKnowledge});
  if(succ.softKnowledge)facts.push({l:'Softskill Knowledge',v:succ.softKnowledge});
  if(succ.calibrationStatus)facts.push({l:'Calibration Status',v:succ.calibrationStatus});
  panel.innerHTML='<div style="background:#f8fafc;border:1px solid var(--border);border-radius:var(--rsm);padding:.8rem 1rem;margin-bottom:4px;">'
    +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:.65rem;flex-wrap:wrap;">'
    +'<span style="font-size:10px;font-weight:700;color:var(--t2);text-transform:uppercase;letter-spacing:.05em;">Learning Status</span>'
    +'<span class="b '+statusCls+'" style="font-size:10px;">'+(learning.status||"—")+'</span>'
    +(learning.overallPct!==null?'<span style="font-size:10px;color:var(--t2);">· Overall <strong>'+learning.overallPct+'%</strong></span>':"")
    +'<span style="font-size:10px;color:var(--t3);margin-left:auto;">Updated: <strong>'+(learning.lastUpdated||"—")+'</strong></span>'
    +'</div>'
    +(facts.length?'<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:.65rem;">'
      +facts.map(f=>'<div style="background:var(--skyxs);border:1px solid var(--border);border-radius:5px;padding:3px 8px;font-size:10px;"><span style="color:var(--t2);">'+_esc(f.l)+':</span> <strong style="color:var(--navy);">'+_esc(String(f.v))+'</strong></div>').join("")
      +'</div>':"")
    +'<div style="font-size:10px;font-weight:700;color:var(--t2);text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px;">Courses &amp; Completion</div>'
    +courseHtml
    +'</div>';
  panel.style.display='block';
}

// ── DETAIL ──
// ── SHARED PAGINATION HELPER ──
const _pgCallbacks={};
function renderPagination(barId, page, totalPages, total, start, end, onPage){
  const bar=document.getElementById(barId);
  if(!bar)return;
  _pgCallbacks[barId]=onPage;
  if(total===0){bar.innerHTML='<span class="pg-info">No records</span>';return;}
  const info='<span class="pg-info">Showing '+(start+1)+'–'+end+' of '+total+'</span>';
  let pages=[];
  if(totalPages<=3){
    for(let i=1;i<=totalPages;i++)pages.push(i);
  } else {
    pages=[1,2];
    if(page>2&&page<totalPages)pages.push(page);
    if(totalPages>2){if(pages[pages.length-1]<totalPages-1)pages.push("…");pages.push(totalPages);}
  }
  // deduplicate
  pages=pages.filter((v,i,a)=>a.indexOf(v)===i);
  const btnPrev='<button class="pg-btn" '+(page<=1?"disabled ":"")+'onclick="_pgGo(\''+barId+'\','+(Math.max(1,page-1))+')">&#8592; Prev</button>';
  const btnNext='<button class="pg-btn" '+(page>=totalPages?"disabled ":"")+'onclick="_pgGo(\''+barId+'\','+(Math.min(totalPages,page+1))+')">Next &#8594;</button>';
  const pgBtns=pages.map(p=>p==="…"
    ?'<span style="padding:4px 6px;color:var(--t3);font-size:11px">…</span>'
    :'<button class="pg-btn'+(p===page?" pg-active":"")+'" onclick="_pgGo(\''+barId+'\','+p+')">'+p+"</button>"
  ).join("");
  bar.innerHTML=info+'<div class="pg-bar-btns">'+btnPrev+pgBtns+btnNext+'</div>';
}
function _pgGo(barId,page){if(_pgCallbacks[barId])_pgCallbacks[barId](page);}

// ── DETAIL PAGINATION STATE ──
let DETAIL_PAGE=1;
const DETAIL_PER_PAGE=10;
let DETAIL_FILTERED=[];

const DCOLS=[
  {h:"Position ID",k:"Position ID"},{h:"Position",k:"Position Name"},{h:"Job Family",k:"Job Family Successor"},{h:"Level",k:"Current Level"},
  {h:"Incumbent",k:"Employee Name"},{h:"Age Group",k:"Group Age"},
  {h:"Retire Risk",k:"Group Age",fn:v=>{const rr=retRiskFromGroup(v);const cl=rr==="Imminent"||rr==="High"?"b-red":rr==="Moderate"?"b-amber":"b-green";return v?'<span class="b '+cl+'">'+ rr+"</span>":"—";}},
  {h:"C-Level",k:"C-Level",fn:v=>isYes(v)?'<span class="b b-gold">Yes</span>':"—"},
  {h:"Critical",k:"Critical Position",fn:v=>isYes(v)?'<span class="b b-red">Yes</span>':"—"},
  {h:"As Talent",k:"As Talent Calibrated",fn:v=>isYes(v)?'<span class="b b-green">YES</span>':(v?'<span class="b b-gray">'+v+"</span>":"—")},
  {h:"Successor",k:"Successor Name",fn:v=>v?'<span style="color:var(--bmed);font-weight:600">'+v+"</span>":"—"},
  {h:"Readiness",k:"Final Score Readiness Text",fn:v=>v?'<span class="b '+rColor(v)+'">'+ v+"</span>":"—"},
  {h:"Business Unit",k:"Business Unit Name"},
];

function detailSearchHandler(){
  DETAIL_PAGE=1;
  renderDetail();
}

function renderDetail(){
  const q=(document.getElementById("detailSearch")||{}).value||"";
  const ql=q.toLowerCase();
  const avail=DCOLS.filter(c=>ALL[0]&&c.k in ALL[0]);
  document.getElementById("detailHead").innerHTML=avail.map(c=>"<th>"+c.h+"</th>").join("");

  const detSrc=tabBuFilter("detailBuFilter","detailClFilter");
  DETAIL_FILTERED=ql
    ? detSrc.filter(row=>
        ["Position Name","Job Family Successor","Employee Name","Successor Name","Business Unit Name","Current Level","Position ID"]
          .some(k=>(row[k]||"").toLowerCase().includes(ql))
      )
    : detSrc.slice();
  // Group same-named incumbents together (and same successor as tiebreaker)
  DETAIL_FILTERED.sort((a,b)=>{
    const an=(a["Employee Name"]||"").toString();
    const bn=(b["Employee Name"]||"").toString();
    const byName=an.localeCompare(bn);
    if(byName!==0)return byName;
    return (a["Position Name"]||"").toString().localeCompare((b["Position Name"]||"").toString());
  });

  const total=DETAIL_FILTERED.length;
  const totalPages=Math.max(1,Math.ceil(total/DETAIL_PER_PAGE));
  if(DETAIL_PAGE>totalPages)DETAIL_PAGE=totalPages;
  const start=(DETAIL_PAGE-1)*DETAIL_PER_PAGE;
  const show=DETAIL_FILTERED.slice(start,start+DETAIL_PER_PAGE);

  const sc=document.getElementById("detailSearchCount");
  if(sc)sc.textContent=ql?total+" result"+(total!==1?"s":"")+" found":"";

  document.getElementById("detailBody").innerHTML=show.length
    ? show.map(row=>"<tr>"+avail.map(c=>{const v=g(row,c.k);const d=c.fn?c.fn(v):v;return"<td>"+(d||"—")+"</td>";}).join("")+"</tr>").join("")
    : '<tr><td colspan="'+avail.length+'" style="text-align:center;color:#94a3b8;font-style:italic;padding:20px">No records match your search.</td></tr>';

  renderPagination("detailPgBar", DETAIL_PAGE, totalPages, total, start, Math.min(start+DETAIL_PER_PAGE,total), (p)=>{DETAIL_PAGE=p;renderDetail();});
}

// ── POSITION PAGINATION STATE ──
let POS_PAGE=1;
const POS_PER_PAGE=6;
let POS_ALL_ARR=[];


// ════════════════════════════════════════════════════
//  SUCCESSION CARD GENERATOR
// ════════════════════════════════════════════════════
function scPopulate(pm){
  const empSel=document.getElementById("scEmp");
  const posSel=document.getElementById("scPos");
  if(!empSel||!posSel)return;
  const pEmp=empSel.value, pPos=posSel.value;

  // unique incumbents
  const people=[...new Set(FILTERED.map(r=>g(r,"Employee Name")).filter(v=>v))].sort();
  empSel.innerHTML='<option value="">— Select an employee —</option>'+people.map(n=>`<option value="${n}">${n}</option>`).join("");
  if(pEmp)empSel.value=pEmp;

  posSel.innerHTML='<option value="">— Select a position —</option>';
  Object.values(pm).sort((a,b)=>a.title.localeCompare(b.title)).forEach(p=>{
    const o=document.createElement("option");
    o.value=p.id;
    o.textContent=p.title+(p.clevel?" [C-Level]":p.critical?" [Critical]":"")+(p.incumbent?" — "+p.incumbent:"");
    posSel.appendChild(o);
  });
  if(pPos)posSel.value=pPos;
}

function scGetData(){
  const empName=document.getElementById("scEmp").value;
  const posId=document.getElementById("scPos").value;
  if(!empName||!posId)return null;

  const pm=buildPM(FILTERED);
  const pos=pm[posId];
  // Incumbent row for this employee
  const eRow=FILTERED.find(r=>g(r,"Employee Name")===empName)||{};
  // Row where this person appears as successor (for readiness etc)
  const sRow=FILTERED.find(r=>g(r,"Successor Name")===empName)||{};

  const dob=g(eRow,"Group Age");
  const age=calcAge(dob);
  const dobDisp=dob?(dob+(age?" (Age "+age+")":"")):"";
  const joinDate=g(eRow,"Join_Date")||g(eRow,"Join_date")||"";
  const yos=g(eRow,"Division")||"";
  const joinDisp=joinDate?(joinDate+(yos?" ("+yos+" yrs)":"")):(yos?yos+" yrs":"");

  return {
    name:empName,
    nik:(()=>{const raw=g(eRow,"Employee_ID")||g(eRow,"NIK")||"";return raw?String(raw).replace(/[^0-9]/g,"").padStart(7,"0"):""})(),
    dob:dobDisp,
    joinDate:joinDisp,
    posTitle:pos?pos.title:"",
    curRole:g(eRow,"Position Name")||"",
    jobFamily:g(eRow,"Job Family Successor")||"",
    level:g(eRow,"Level Successor")||"",
    supervisor:g(eRow,"Current_Supervisor")||g(eRow,"Supervisor")||"",
    bu:g(eRow,"BU")||g(eRow,"Business_Unit")||g(eRow,"BU_CC")||"",
    edu:g(eRow,"Education")||"",
    readiness:g(sRow,"Final Score Readiness Text")||"",
    pat2024:g(eRow,"PAT_2024")||g(eRow,"PAT2024")||"",
    pat2023:g(eRow,"PAT_2023")||g(eRow,"PAT2023")||"",
    pat2022:g(eRow,"PAT_2022")||g(eRow,"PAT2022")||"",
    feedScore:g(eRow,"Score_360")||g(eRow,"Feedback_Score")||"",
    feedText:g(eRow,"Feedback_360")||g(eRow,"Feedback")||"",
    aspiration:g(eRow,"Personal_Aspiration")||g(eRow,"Aspiration")||"",
    certif:g(eRow,"Certification")||"",
    careerIn:g(eRow,"Career_History_Internal")||g(eRow,"Career_Internal")||"",
    careerOut:g(eRow,"Career_History_External")||g(eRow,"Career_External")||"",
    supStrength:g(eRow,"Supervisor_Strength")||g(eRow,"Strength")||"",
    supDev:g(eRow,"Supervisor_Development")||g(eRow,"Development")||"",
    overallInsight:g(eRow,"Overall_Insight")||"",
    readinessWithin1Y:g(eRow,"Readiness_Within_1Y")||"",
    readiness1to3Y:g(eRow,"Readiness_1_3Y")||"",
    dev70:g(eRow,"Dev_70_Experience")||"",
    dev20:g(eRow,"Dev_20_Exposure")||"",
    dev10:g(eRow,"Dev_10_Education")||"",
  };
}

function scPreview(){
  const d=scGetData();
  const box=document.getElementById("scPreviewBox");
  if(!d){box.style.display="none";return;}
  box.style.display="block";
  const fields=[
    ["Employee",d.name],["NIK",d.nik],["Targeted Position",d.posTitle],
    ["Date of Birth / Age",d.dob],["Join Date / YoS",d.joinDate],
    ["Current Role",d.curRole],["Job Family",d.jobFamily],["Level",d.level],
    ["Current Supervisor",d.supervisor],["BU / CC Group",d.bu],
    ["Education",d.edu],["Readiness",d.readiness],
  ];
  document.getElementById("scPreviewGrid").innerHTML=fields.map(([lbl,val])=>`
    <div>
      <div style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--t3);margin-bottom:2px">${lbl}</div>
      <div style="font-size:12px;font-weight:600;color:var(--navy)">${val||'<span style="color:#94a3b8;font-weight:400">—</span>'}</div>
    </div>`).join("");
  document.getElementById("scStatusMsg").textContent="";
}

function scGenerate(){
  const d=scGetData();
  if(!d){
    alert("Please select both an employee and a targeted position.");
    return;
  }
  const btn=document.getElementById("scGenBtn");
  btn.disabled=true;btn.textContent="Building...";
  document.getElementById("scStatusMsg").style.color="var(--bmed)";
  document.getElementById("scStatusMsg").textContent="Generating PPTX — please wait...";

  // Dynamically load JSZip then build
  function loadScript(src){
    return new Promise((res,rej)=>{
      if(document.querySelector('script[src="'+src+'"]')){res();return;}
      const s=document.createElement("script");s.src=src;s.onload=res;s.onerror=rej;document.head.appendChild(s);
    });
  }
  loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js")
    .then(()=>scBuildPPTX(d))
    .then(blob=>{
      const url=URL.createObjectURL(blob);
      const a=document.createElement("a");
      a.href=url;
      a.download="SuccessionCard_"+d.name.replace(/\s+/g,"_")+"_to_"+d.posTitle.replace(/\s+/g,"_")+".pptx";
      document.body.appendChild(a);a.click();
      document.body.removeChild(a);URL.revokeObjectURL(url);
      btn.disabled=false;btn.textContent="Generate Card";
      document.getElementById("scStatusMsg").style.color="var(--green)";
      document.getElementById("scStatusMsg").textContent="✓ Succession Card downloaded successfully!";
    })
    .catch(err=>{
      btn.disabled=false;btn.textContent="Generate Card";
      document.getElementById("scStatusMsg").style.color="var(--red)";
      document.getElementById("scStatusMsg").textContent="✗ Error: "+err.message;
      console.error(err);
    });
}

// ── Core PPTX builder ──────────────────────────────────────────────
async function scBuildPPTX(d){
  const JSZip=window.JSZip;
  if(!JSZip)throw new Error("JSZip not loaded");

  const TEMPLATE_B64="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cHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0MS54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAAB52AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0Mi54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAACR3AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0My54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAACp4AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0NS54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAADB5AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0Ni54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAADZ6AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0Ny54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAADx7AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0OC54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAsAAAAAAAAAAAAAAAAAEJ8AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0NC54bWwucmVsc1BLAQItABQABgAIAAAAIQDV0ZLxvAAAADcBAAAtAAAAAAAAAAAAAAAAAEh9AABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRlTGF5b3V0MTAueG1sLnJlbHNQSwECLQAUAAYACAAAACEA1dGS8bwAAAA3AQAALQAAAAAAAAAAAAAAAABPfgAAcHB0L3NsaWRlTGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDExLnhtbC5yZWxzUEsBAi0AFAAGAAgAAAAhAEEwquTzBgAAYmEAACEAAAAAAAAAAAAAAAAAVn8AAHBwdC9ub3Rlc01hc3RlcnMvbm90ZXNNYXN0ZXIxLnhtbFBLAQItABQABgAIAAAAIQCS5wuNxwUAAEcbAAAUAAAAAAAAAAAAAAAAAIiGAABwcHQvdGhlbWUvdGhlbWUxLnhtbFBLAQItABQABgAIAAAAIQC0z1gZuQAAACQBAAAsAAAAAAAAAAAAAAAAAIGMAABwcHQvbm90ZXNNYXN0ZXJzL19yZWxzL25vdGVzTWFzdGVyMS54bWwucmVsc1BLAQItABQABgAIAAAAIQCS5wuNxwUAAEcbAAAUAAAAAAAAAAAAAAAAAISNAABwcHQvdGhlbWUvdGhlbWUyLnhtbFBLAQItAAoAAAAAAAAAIQDH6D0OeQ0AAHkNAAAUAAAAAAAAAAAAAAAAAH2TAABwcHQvbWVkaWEvaW1hZ2UxLnBuZ1BLAQItABQABgAIAAAAIQBoJNHIHQIAAAMFAAARAAAAAAAAAAAAAAAAACihAABwcHQvcHJlc1Byb3BzLnhtbFBLAQItABQABgAIAAAAIQAjAczabQEAAAIDAAARAAAAAAAAAAAAAAAAAHSjAABwcHQvdmlld1Byb3BzLnhtbFBLAQItABQABgAIAAAAIQBshiEQdQEAACkFAAATAAAAAAAAAAAAAAAAABClAABwcHQvdGFibGVTdHlsZXMueG1sUEsBAi0AFAAGAAgAAAAhAOrD5AP7AQAA4QQAABAAAAAAAAAAAAAAAAAAtqYAAGRvY1Byb3BzL2FwcC54bWxQSwECLQAUAAYACAAAACEASZTuR1kBAACQAgAAEQAAAAAAAAAAAAAAAADnqQAAZG9jUHJvcHMvY29yZS54bWxQSwECLQAUAAYACAAAACEAYsHgM6wCAACCBwAADAAAAAAAAAAAAAAAAAB3rAAAcHB0L21ldGFkYXRhUEsFBgAAAAAtAC0Avw0AAE2vAAAAAA==";

  function b64ToUint8(s){
    const bin=atob(s);const arr=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);return arr;
  }
  const zip=await JSZip.loadAsync(b64ToUint8(TEMPLATE_B64));
  const xe=v=>String(v||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

  // Build a run of text with Calibri 9pt black styling
  function makeRun(value){
    return `<a:r><a:rPr lang="en-US" sz="900" b="0" i="0" u="none" strike="noStrike" cap="none"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/></a:rPr><a:t>${xe(value)}</a:t></a:r>`;
  }

  // Inject text into a <a:tc> block — replaces paragraph content, keeps cell fill/borders
  function injectTc(tcXml, value){
    if(!value) return tcXml;
    // Keep pPr if it exists, replace everything else inside <a:p>
    const pPrMatch = tcXml.match(/<a:pPr[^>]*\/>/);
    const pPr = pPrMatch ? pPrMatch[0] : '';
    return tcXml.replace(
      /(<a:txBody>[\s\S]*?<a:p>)[\s\S]*?(<\/a:p>[\s\S]*?<\/a:txBody>)/,
      (_, before, after) => `${before}${pPr}${makeRun(value)}${after}`
    );
  }

  let slideXml = await zip.file("ppt/slides/slide1.xml").async("string");

  // --- Title shape: "Succession Card for Targeted Position - Name - NIK" ---
  const titleVal = xe("Succession Card for Targeted Position - " + (d.name||"Unknown") + " - " + (d.nik||"—"));
  slideXml = slideXml.replace(/<a:t>XXX - xxx<\/a:t>/, `<a:t>${titleVal}</a:t>`);

  // --- Table cells: split into array, inject by index ---
  // Value cells (white fields next to grey labels):
  // [1]=Current Role, [3]=Job Family, [5]=Level, [7]=Supervisor, [9]=BU
  // [11]=DoB/Age, [13]=Join Date/YoS, [15]=PAT, [17]=Education
  // [59]=Personal Aspiration, [61]=360 Feedback
  // Targeted Position is a label in [40]; its value goes in [41] (empty white cell)
  const valueCells = {
    1:  d.curRole    || "",
    3:  d.jobFamily  || "",
    5:  d.level      || "",
    7:  d.supervisor || "",
    9:  d.bu         || "",
    11: d.dob        || "",
    13: d.joinDate   || "",
    17: d.edu        || "",
    41: d.posTitle   || "",
    59: d.aspiration || "",
    61: d.feedback   || "",
  };

  // Parse all <a:tc> blocks with their positions
  const tcRegex = /<a:tc>[\s\S]*?<\/a:tc>/g;
  const allTcs = [];
  let m;
  while((m = tcRegex.exec(slideXml)) !== null)
    allTcs.push({start: m.index, end: m.index + m[0].length, text: m[0]});

  // Inject in reverse order to preserve string indices
  const indices = Object.keys(valueCells).map(Number).sort((a,b)=>b-a);
  for(const idx of indices){
    const val = valueCells[idx];
    if(!val || !allTcs[idx]) continue;
    const tc = allTcs[idx];
    const replaced = injectTc(tc.text, val);
    slideXml = slideXml.slice(0, tc.start) + replaced + slideXml.slice(tc.end);
  }

  zip.file("ppt/slides/slide1.xml", slideXml);

  return await zip.generateAsync({
    type:"blob",
    mimeType:"application/vnd.openxmlformats-officedocument.presentationml.presentation"
  });
}


// ════════════════════════════════════════════════════
//  EXECUTIVE SUMMARY
// ════════════════════════════════════════════════════
function devRecommendation(p,bd){
  const rs=riskScore(p);
  if(p.successors.length===0&&p.critical)return{icon:"🚨",bg:"#FEE2E2",title:"External Hiring Recommended",sub:p.title+" has no successor and is critical."};
  if(p.successors.length===0&&p.clevel)return{icon:"🚨",bg:"#FEE2E2",title:"Urgent: Build C-Level Pipeline",sub:"No successor identified for C-Level role: "+p.title};
  if(bd.now===0&&p.clevel)return{icon:"⚡",bg:"#FEF3C7",title:"Accelerate "+p.title+" Successor",sub:"No ready-now candidate. Fastest available: "+(bd.y1?"1-3 years":"3-5 years")+"."};
  if(p.flightRisk.toLowerCase()==="high"&&p.successors.length<2)return{icon:"⚠️",bg:"#FEF3C7",title:"Retention Risk — "+p.title,sub:"High flight risk incumbent with thin successor bench."};
  if(p.retRisk==="Imminent"&&bd.now===0)return{icon:"⏰",bg:"#FEF3C7",title:"Retirement Imminent — Accelerate Readiness",sub:p.title+" incumbent retiring soon, no ready-now successor."};
  if(p.successors.length>=1&&bd.now>0&&rs<30)return{icon:"✅",bg:"#D1FAE5",title:"Maintain Momentum — "+p.title,sub:"Bench is healthy. Continue development plan."};
  return null;
}

// ── EXEC SUMMARY PAGINATION STATE ──
const ES={
  risks:{page:1,per:8},
  recs:{page:1,per:5},
  load:{page:1,per:10},
  tgap:{page:1,per:10},
  bench:{page:1,per:10},
};
// Store full sorted arrays so pagination can slice them
let ES_DATA={risks:[],recs:[],load:[],tgap:[],bench:[]};

function esPage(key,p){ES[key].page=p;renderESSection(key);}

function renderESSection(key){
  const s=ES[key];
  const data=ES_DATA[key];
  const total=data.length;
  const totalPages=Math.max(1,Math.ceil(total/s.per));
  if(s.page>totalPages)s.page=totalPages;
  const start=(s.page-1)*s.per;
  const slice=data.slice(start,start+s.per);

  if(key==="risks"){
    const tb=document.getElementById("topRisksBody");
    if(tb)tb.innerHTML=slice.map(p=>{
      const rs=riskScore(p);const rl=riskLabel(rs);
      const issues=[];
      if(p.successors.length===0)issues.push("No successor");
      if(p.retRisk==="High")issues.push("Retire: High");
      if(p.flightRisk.toLowerCase()==="high")issues.push("High flight risk");
      if(p.critical)issues.push("Critical role");
      return`<tr>
        <td style="font-weight:600;color:var(--navy)">${p.title}</td>
        <td><span style="font-weight:700;color:${rs>=60?"#EF4444":rs>=30?"#F59E0B":"#10B981"}">${rs}/100</span></td>
        <td><span class="risk-badge ${rl.cls}"><span class="risk-dot ${rl.dot}"></span>${rl.txt}</span></td>
        <td style="font-size:10px;color:var(--t2)">${issues.join(" · ")||"—"}</td>
      </tr>`;
    }).join("")||'<tr><td colspan="4" style="color:#94a3b8;font-style:italic;padding:10px">No data.</td></tr>';
    renderPagination("pgTopRisks",s.page,totalPages,total,start,Math.min(start+s.per,total),p=>esPage("risks",p));
  }

  if(key==="recs"){
    const devEl=document.getElementById("devRecsBody");
    if(devEl)devEl.innerHTML=slice.length?slice.map(r=>`
      <div class="dev-rec" style="background:${r.bg}22;border-color:${r.bg}88">
        <div class="dev-rec-icon" style="background:${r.bg}">${r.icon}</div>
        <div class="dev-rec-body">
          <div class="dev-rec-title">${r.title}</div>
          <div class="dev-rec-sub">${r.sub}</div>
        </div>
      </div>`).join(""):'<div style="color:var(--t2);font-size:12px;padding:.5rem">No urgent recommendations — succession posture is healthy.</div>';
    renderPagination("pgDevRecs",s.page,totalPages,total,start,Math.min(start+s.per,total),p=>esPage("recs",p));
  }

  if(key==="load"){
    const loadEl=document.getElementById("succLoadBody");
    if(loadEl)loadEl.innerHTML=slice.map(({name,count})=>{
      const cls=count>=4?"load-over":count>=2?"load-warn":"load-ok";
      const icon=count>=4?"🔴":count>=2?"⚠️":"✅";
      return`<tr>
        <td style="font-weight:600;color:var(--navy)">${name}</td>
        <td style="font-weight:700">${count} role${count!==1?"s":""}</td>
        <td><span class="load-badge ${cls}">${icon} ${count>=4?"Overloaded":count>=2?"Watch":"OK"}</span></td>
      </tr>`;
    }).join("")||'<tr><td colspan="3" style="color:#94a3b8;font-style:italic;padding:10px">No data.</td></tr>';
    renderPagination("pgSuccLoad",s.page,totalPages,total,start,Math.min(start+s.per,total),p=>esPage("load",p));
  }

  if(key==="tgap"){
    const tgEl=document.getElementById("timeGapBody");
    if(tgEl)tgEl.innerHTML=slice.map(p=>{
      const best=bestReadiness(p);
      return`<tr>
        <td style="font-weight:600;color:var(--navy)">${p.title}</td>
        <td style="font-size:11px">${best==="none"?"—":best}</td>
        <td>${timeGapHtml(best)}</td>
      </tr>`;
    }).join("")||'<tr><td colspan="3" style="color:#94a3b8;font-style:italic;padding:10px">No data.</td></tr>';
    renderPagination("pgTimeGap",s.page,totalPages,total,start,Math.min(start+s.per,total),p=>esPage("tgap",p));
  }

  if(key==="bench"){
    const bdEl=document.getElementById("benchDepthBody");
    if(bdEl)bdEl.innerHTML=slice.map(p=>{
      const rs=riskScore(p);const rl=riskLabel(rs);
      const bd=benchDepth(p);
      const bdHtml='<div class="bench-depth">'
        +(bd.now?'<span class="bench-tier bt-now">Now: '+bd.now+"</span>":"")
        +(bd.y1?'<span class="bench-tier bt-1y">1-3yr: '+bd.y1+"</span>":"")
        +(bd.y3p?'<span class="bench-tier bt-3p">3-5yr: '+bd.y3p+"</span>":"")
        +(p.successors.length===0?'<span class="bench-tier" style="background:#FEE2E2;color:#991b1b">None</span>':"")
        +"</div>";
      const fr=(p.flightRisk||"").toLowerCase();
      const frBadge=fr==="high"?'<span class="b b-red">High</span>':fr==="medium"||fr==="moderate"?'<span class="b b-amber">Medium</span>':fr&&fr!=="unknown"?'<span class="b b-green">'+p.flightRisk+'</span>':'<span style="color:var(--t3);font-size:11px">—</span>';
      return`<tr>
        <td style="font-weight:600;color:var(--navy)">${p.title}</td>
        <td>${p.clevel?'<span class="b b-gold">Yes</span>':"—"}</td>
        <td>${p.critical?'<span class="b b-red">Yes</span>':"—"}</td>
        <td>${frBadge}</td>
        <td>${bdHtml}</td>
      </tr>`;
    }).join("")||'<tr><td colspan="5" style="color:#94a3b8;font-style:italic;padding:10px">No data.</td></tr>';
    renderPagination("pgBenchDepth",s.page,totalPages,total,start,Math.min(start+s.per,total),p=>esPage("bench",p));
  }
}

function renderExecSummary(){
  const pm=buildPM(FILTERED);
  const arr=Object.values(pm);
  const total=arr.length||1;

  // Health Score
  const scores=arr.map(p=>riskScore(p));
  const avgRisk=scores.reduce((a,b)=>a+b,0)/total;
  const healthPct=Math.round(100-avgRisk);
  const healthEl=document.getElementById("healthPct");
  const healthBar=document.getElementById("healthBarFill");
  const healthLbl=document.getElementById("healthLabel");
  if(healthEl){
    healthEl.textContent=healthPct+"%";
    healthEl.style.color=healthPct>=70?"#10B981":healthPct>=45?"#F59E0B":"#EF4444";
    healthBar.style.width=healthPct+"%";
    healthBar.style.background=healthPct>=70?"linear-gradient(90deg,#10B981,#3B82F6)":healthPct>=45?"linear-gradient(90deg,#F59E0B,#3B82F6)":"linear-gradient(90deg,#EF4444,#F59E0B)";
    healthLbl.textContent=healthPct>=70?"Overall succession posture is healthy.":healthPct>=45?"Succession posture needs attention — review Watch and Critical positions.":"Succession posture is at risk — immediate action required on Critical positions.";
  }

  // KPIs
  const criticalCount=arr.filter(p=>riskScore(p)>=60).length;
  const noSucc=arr.filter(p=>p.successors.length===0).length;
  const clArr=arr.filter(p=>p.clevel);
  const clCoverage=clArr.length?Math.round(clArr.filter(p=>p.successors.length>0).length/clArr.length*100):0;
  const readyNowCount=arr.filter(p=>p.successors.some(s=>(s.readiness||"").toLowerCase().includes("now"))).length;
  const readyNowPct=Math.round(readyNowCount/total*100);

  const kpis=[
    //{val:healthPct+"%",lbl:"Health Score",sub:"Overall posture",color:healthPct>=70?"#10B981":healthPct>=45?"#F59E0B":"#EF4444"},
    {val:criticalCount,lbl:"Critical Risks",sub:"Positions at critical risk",color:"#EF4444"},
    {val:noSucc,lbl:"No Successor",sub:"Roles with zero bench",color:"#EF4444"},
    {val:clCoverage+"%",lbl:"CXO Coverage",sub:"C-Level positions covered",color:"#1E6FD9"},
    {val:readyNowPct+"%",lbl:"Ready-Now Coverage",sub:"Positions with ready-now successor",color:"#10B981"},
  ];
  const grid=document.getElementById("execKpiGrid");
  if(grid)grid.innerHTML=kpis.map(k=>`
    <div class="exec-kpi">
      <div class="exec-kpi-val" style="color:${k.color}">${k.val}</div>
      <div class="exec-kpi-lbl">${k.lbl}</div>
      <div class="exec-kpi-sub">${k.sub}</div>
    </div>`).join("");

  // Build full sorted data arrays
  const sortedByRisk=[...arr].sort((a,b)=>riskScore(b)-riskScore(a));

  ES_DATA.risks=sortedByRisk;

  const allRecs=[];
  for(const p of sortedByRisk){
    const bd=benchDepth(p);
    const rec=devRecommendation(p,bd);
    if(rec)allRecs.push(rec);
  }
  ES_DATA.recs=allRecs;

  const loadMap={};
  FILTERED.forEach(r=>{
    const sn=g(r,"Successor Name");if(!sn)return;
    if(!loadMap[sn])loadMap[sn]=new Set();
    const pid=g(r,"Position ID");if(pid)loadMap[sn].add(pid);
  });
  ES_DATA.load=Object.entries(loadMap).map(([name,pids])=>({name,count:pids.size})).sort((a,b)=>b.count-a.count);

  ES_DATA.tgap=sortedByRisk;
  ES_DATA.bench=sortedByRisk;

  // Reset pages to 1 on fresh render
  Object.keys(ES).forEach(k=>{ES[k].page=1;});

  // Render all sections
  ["risks","recs","load","tgap","bench"].forEach(k=>renderESSection(k));
}


function openPosDetail(pid){
  const pm=buildPM(FILTERED);
  const p=pm[pid];if(!p)return;
  document.getElementById("pdTitle").textContent=p.title;
  document.getElementById("pdIncumbent").innerHTML=p.incumbent
    ?'<svg viewBox="0 0 24 24" fill="none" stroke="var(--t3)" stroke-width="2" stroke-linecap="round" style="width:11px;height:11px;vertical-align:middle;margin-right:4px"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>'+p.incumbent
    :"No incumbent on record";

  const rs=riskScore(p);const rl=riskLabel(rs);
  const bd=benchDepth(p);
  const best=bestReadiness(p);

  const tag=(txt,cls)=>'<span class="b '+cls+'" style="margin-right:3px;margin-bottom:3px">'+txt+'</span>';
  const tagRow=[];
  if(p.clevel)tagRow.push(tag("C-Level","b-gold"));
  if(p.critical)tagRow.push(tag("Critical","b-red"));
  const fr=(p.flightRisk||"").toLowerCase();
  if(fr&&fr!=="unknown")tagRow.push(tag("Flight Risk: "+p.flightRisk,fr==="high"?"b-red":fr==="medium"?"b-amber":"b-green"));
  if(p.retRisk==="High")tagRow.push(tag("Retire: High","b-red"));

  const row=(lbl,val)=>val?`<div class="pos-detail-row"><span class="pos-detail-lbl">${lbl}</span><span class="pos-detail-val">${val}</span></div>`:"";
  const fmtDate=d=>{
    if(d===null||d===undefined||d==="")return"";
    // Real Date object
    if(d instanceof Date)return d.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"});
    // Excel serial date number (days since 1900-01-01, with Excel's leap-year bug accounted for)
    // SheetJS returns these when a cell is stored as a number rather than a date type.
    const n=typeof d==="number"?d:(typeof d==="string"&&/^\d+(\.\d+)?$/.test(d.trim())?parseFloat(d):NaN);
    if(!isNaN(n)&&n>20000&&n<60000){ // plausible date-serial range (~1954 to ~2064)
      const ms=Math.round((n-25569)*86400*1000);
      const dt=new Date(ms);
      if(!isNaN(dt.getTime()))return dt.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"});
    }
    return String(d);
  };
  const skillList=(arr)=>arr.length?arr.map(s=>`<div style="display:flex;justify-content:space-between;font-size:11px;padding:3px 0;border-bottom:1px solid var(--border)"><span style="color:var(--navy);font-weight:500">${s.desc}</span><span style="color:var(--t2);font-size:10px">${s.threshold||""}</span></div>`).join(""):"—";

  const succRows=p.successors.length
    ?p.successors.map(s=>`
      <div style="background:var(--skyxs);border:1px solid var(--border);border-radius:8px;padding:.85rem 1rem;margin-bottom:.7rem;">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:4px;">
          <span style="font-size:13px;font-weight:700;color:var(--navy)">${s.name}</span>
          <div style="display:flex;gap:4px;flex-wrap:wrap;">
            ${s.asTalent?'<span class="b b-green" style="font-size:9px">Talent</span>':""}
            ${s.calibrationStatus?'<span class="b b-blue" style="font-size:9px">'+s.calibrationStatus+'</span>':""}
            ${s.readiness?'<span class="b '+rColor(s.readiness)+'" style="font-size:9px">'+s.readiness+'</span>':""}
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:3px 16px;font-size:11px;margin-bottom:8px;">
          ${s.succEmp?'<div><span style="color:var(--t2)">NIK: </span><b style="color:var(--navy)">'+s.succEmp+'</b></div>':""}
          ${s.level?'<div><span style="color:var(--t2)">Level: </span><b style="color:var(--navy)">'+s.level+'</b></div>':""}
          ${s.succPosition?'<div style="grid-column:1/-1"><span style="color:var(--t2)">Current Position: </span><b style="color:var(--navy)">'+s.succPosition+'</b></div>':""}
          ${s.family?'<div><span style="color:var(--t2)">Job Family: </span><b style="color:var(--navy)">'+s.family+'</b></div>':""}
          ${s.bu?'<div><span style="color:var(--t2)">Business Unit: </span><b style="color:var(--navy)">'+s.bu+'</b></div>':""}
          ${s.flightRisk?'<div><span style="color:var(--t2)">Flight Risk: </span><b style="color:var(--navy)">'+s.flightRisk+'</b></div>':""}
        </div>
        ${(s.coach||s.mentor||s.mentorNom1||s.mentorNom2||s.projectAssignment||s.certPlan||s.trainingName)?`
        <div style="margin-top:8px;padding-top:8px;border-top:1px solid var(--border);display:grid;grid-template-columns:1fr 1fr;gap:3px 16px;font-size:11px;">
          ${s.coach?'<div><span style="color:var(--t2)">Coach: </span><b style="color:var(--navy)">'+s.coach+'</b></div>':""}
          ${s.mentor?'<div><span style="color:var(--t2)">Mentor: </span><b style="color:var(--navy)">'+s.mentor+'</b></div>':""}
          ${s.mentorNom1?'<div><span style="color:var(--t2)">Mentor Nom. 1: </span><b style="color:var(--navy)">'+s.mentorNom1+'</b></div>':""}
          ${s.mentorNom2?'<div><span style="color:var(--t2)">Mentor Nom. 2: </span><b style="color:var(--navy)">'+s.mentorNom2+'</b></div>':""}
          ${s.projectAssignment?'<div style="grid-column:1/-1"><span style="color:var(--t2)">Project Assignment: </span><b style="color:var(--navy)">'+s.projectAssignment+'</b></div>':""}
          ${s.certPlan?'<div style="grid-column:1/-1"><span style="color:var(--t2)">Certification Plan: </span><b style="color:var(--navy)">'+s.certPlan+'</b></div>':""}
          ${s.trainingName?'<div style="grid-column:1/-1"><span style="color:var(--t2)">Training: </span><b style="color:var(--navy)">'+s.trainingName+(s.trainingPlan?' ('+s.trainingPlan+')':'')+'</b></div>':""}
        </div>`:""}
      </div>`).join("")
    :'<div class="nosucc">No successor identified</div>';

  document.getElementById("posDetailBody").innerHTML=`
    <div class="pos-detail-section">
      <div class="pos-detail-section-title">Position Info</div>
      ${row("NIK (Incumbent)",p.nik)}
      ${row("Position Status",p.positionStatus)}
      ${row("Pillar",p.pillarName)}
      ${row("Job Family",p.family)}
      ${row("Current Level",p.level)}
      ${row("Start Date",fmtDate(p.startDate))}
      ${row("Last Promotion",fmtDate(p.lastPromotion))}
      ${row("Last Rotation",fmtDate(p.lastRotation))}
    </div>
    <div class="pos-detail-section">
      <div class="pos-detail-section-title">Organisation</div>
      ${row("Business Unit",p.bu)}
      ${row("Group",p.group)}
      ${row("Directorate",p.directorate)}
      ${row("Department",p.department)}
      ${row("Division",p.division)}
      ${row("Supervisor",p.supervisor)}
      ${row("Supervisor Position",p.spvPositionName)}
    </div>
    <div class="pos-detail-section">
      <div class="pos-detail-section-title">Risk & Status</div>
      <div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px">${tagRow.join("")||'<span style="font-size:11px;color:var(--t3)">No tags</span>'}</div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">${timeGapHtml(best)}<span class="risk-badge ${rl.cls}"><span class="risk-dot ${rl.dot}"></span>${rl.txt}</span></div>
    </div>
    <div class="pos-detail-section">
      <div class="pos-detail-section-title">Bench Depth</div>
      <div class="bench-depth">
        ${bd.now?'<span class="bench-tier bt-now">✓ Ready Now: '+bd.now+'</span>':""}
        ${bd.y1?'<span class="bench-tier bt-1y">1–3 yrs: '+bd.y1+'</span>':""}
        ${bd.y3p?'<span class="bench-tier bt-3p">3–5 yrs: '+bd.y3p+'</span>':""}
        ${p.successors.length===0?'<span class="bench-tier" style="background:#FEE2E2;color:#991b1b">No Bench</span>':""}
      </div>
    </div>
    <div class="pos-detail-section">
      <div class="pos-detail-section-title">Successors (${p.successors.length})</div>
      ${succRows}
    </div>`;
  document.getElementById("posDetailBackdrop").classList.add("show");
  document.body.style.overflow="hidden";
}
function closePosDetail(){
  document.getElementById("posDetailBackdrop").classList.remove("show");
  document.body.style.overflow="";
}
document.addEventListener("keydown",e=>{if(e.key==="Escape")closePosDetail();});
function exportPDF(){
  const l=document.getElementById("pdfLoading");l.classList.add("show");
  setTimeout(()=>{try{window.print();}catch(e){}l.classList.remove("show");},200);
}

// ═══════════════════════════════════════════════════════════════════════════
// FEATURE 1: READINESS DETAIL MODAL
// ═══════════════════════════════════════════════════════════════════════════
// Per-card "View Details" opens a searchable, filterable table of all
// successors in that readiness bucket, joined with MyLearning data when
// available. State lives in _rdState; rendering is incremental.
// ═══════════════════════════════════════════════════════════════════════════
const _rdState = { category: null, rows: [] };

function _categoryFromCls(cls){
  if (cls === "rn") return "Ready Now";
  if (cls === "rl") return "Ready Later";
  if (cls === "rf") return "Ready Future";
  return null;
}
function _matchesReadiness(readinessText, category){
  const r = (readinessText||"").toLowerCase().trim();
  if (category === "Ready Now")    return r.includes("ready now") || r.includes("less than a year");
  if (category === "Ready Later")  return r.includes("ready later") || r.includes("1-3");
  if (category === "Ready Future") return r.includes("ready future") || r.includes("3-5") || r.includes("3+");
  return false;
}

// Build the successor list for a readiness category, joining with MyLearning
function _buildReadinessRows(category){
  const seen = new Set();
  const rows = [];
  (FILTERED.length ? FILTERED : ALL).forEach(r => {
    const succName = g(r, "Successor Name");
    if (!succName) return;
    const readiness = g(r, "Final Score Readiness Text");
    if (!_matchesReadiness(readiness, category)) return;
    const succNik = g(r, "Succession Emp Number");
    const dedupeKey = (succNik || succName).toLowerCase().trim() + "|" + (g(r,"Position ID")||g(r,"Position Name")||"");
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    const learning = _learningForEmployee(succNik, succName);
    rows.push({
      succNik,
      succName,
      successorPosition: g(r, "Successor Position") || g(r, "Job Family Successor") || "—",
      successorForPos: g(r, "Position Name") || "—",
      department: g(r, "Departement Name") || g(r, "Division") || "—",
      businessUnit: g(r, "Business Unit Name") || "—",
      isCLevel: isYes(g(r, "C-Level")),
      readiness: readiness || "—",
      readinessCls: category === "Ready Now" ? "rn" : category === "Ready Later" ? "rl" : "rf",
      courses: learning.courses,
      overallPct: learning.overallPct,
      learningStatus: learning.status,
      lastUpdated: learning.lastUpdated || g(r,"Last Promotion Date") || g(r,"Start Date") || "—"
    });
  });
  return rows;
}

// Look up learning data for an employee (NIK preferred, name fallback)
function _learningForEmployee(nik, name){
  const s = SOURCES.mylearning;
  if (!s || !s.raw.length || !s.mapping || Object.keys(s.mapping).length === 0) {
    return { courses: [], overallPct: null, status: "—", lastUpdated: null };
  }
  const matchRows = [];
  const nNik = nik ? _norm(nik) : null;
  const nName = name ? _norm(name) : null;
  s.raw.forEach(row => {
    const rowNik = _extractCanonValue(row, s.mapping, "NIK");
    const rowName = _extractCanonValue(row, s.mapping, "Employee Name");
    if (nNik && rowNik && _norm(rowNik) === nNik) { matchRows.push(row); return; }
    if (nName && rowName && _norm(rowName) === nName) matchRows.push(row);
  });
  if (!matchRows.length) return { courses: [], overallPct: null, status: "Not Started", lastUpdated: null };

  const courses = [];
  const pcts = [];
  let lastUpdated = null;
  matchRows.forEach(row => {
    const courseName = _extractCanonValue(row, s.mapping, "Training Name");
    if (!courseName) return;
    const pctRaw = _extractCanonValue(row, s.mapping, "Completion Percentage");
    const pct = pctRaw ? parseFloat(String(pctRaw).replace(/%/g,"")) : null;
    const stat = _extractCanonValue(row, s.mapping, "Training Status");
    const upd = _extractCanonValue(row, s.mapping, "Training Date");
    if (upd && (!lastUpdated || String(upd) > String(lastUpdated))) lastUpdated = upd;
    let courseStatus = stat || (pct >= 100 ? "Completed" : pct > 0 ? "In Progress" : "Not Started");
    courses.push({ name: courseName, pct: isNaN(pct) ? null : pct, status: courseStatus });
    if (!isNaN(pct) && pct !== null) pcts.push(pct);
  });
  const overallPct = pcts.length ? Math.round(pcts.reduce((a,b)=>a+b,0) / pcts.length) : null;
  // Status rollup: Completed if all 100, Not Started if all 0, else In Progress
  let status = "—";
  if (pcts.length) {
    if (pcts.every(p => p >= 100)) status = "Completed";
    else if (pcts.every(p => p === 0)) status = "Not Started";
    else status = "In Progress";
  }
  return { courses, overallPct, status, lastUpdated };
}

function openReadinessDetail(cls){
  const category = _categoryFromCls(cls);
  if (!category) return;
  _rdState.category = category;
  _rdState.rows = _buildReadinessRows(category);

  document.getElementById("rdTitle").textContent = category + " · Successor Details";
  document.getElementById("rdSub").textContent =
    `${_rdState.rows.length} successors • Click any row to inspect their position record`;

  // Populate department filter
  const deptSet = new Set();
  _rdState.rows.forEach(r => { if (r.department && r.department !== "—") deptSet.add(r.department); });
  const deptSel = document.getElementById("rdDept");
  deptSel.innerHTML = '<option value="">All Departments</option>' +
    Array.from(deptSet).sort().map(d => `<option value="${_esc(d)}">${_esc(d)}</option>`).join("");

  // Reset filters
  document.getElementById("rdSearch").value = "";
  document.getElementById("rdLevel").value = "";
  document.getElementById("rdLearnStatus").value = "";
  document.getElementById("rdPct").value = "";

  renderReadinessDetailRows();
  document.getElementById("rdBackdrop").classList.add("show");
  document.body.style.overflow = "hidden";
}

function closeReadinessDetail(){
  document.getElementById("rdBackdrop").classList.remove("show");
  document.body.style.overflow = "";
}

function renderReadinessDetailRows(){
  const q = (document.getElementById("rdSearch").value || "").toLowerCase().trim();
  const dept = document.getElementById("rdDept").value;
  const level = document.getElementById("rdLevel").value;
  const learnStatus = document.getElementById("rdLearnStatus").value;
  const pctThresh = document.getElementById("rdPct").value;

  const filtered = _rdState.rows.filter(r => {
    if (q && !(r.succName.toLowerCase().includes(q) || (r.succNik||"").toLowerCase().includes(q))) return false;
    if (dept && r.department !== dept) return false;
    if (level === "c" && !r.isCLevel) return false;
    if (level === "non" && r.isCLevel) return false;
    if (learnStatus && r.learningStatus !== learnStatus) return false;
    if (pctThresh !== "") {
      const t = parseFloat(pctThresh);
      if (r.overallPct === null) return t === 0;  // null counts as "Not Started" for 0% filter
      if (t === 100) return r.overallPct >= 100;
      if (t === 0) return r.overallPct === 0;
      if (r.overallPct < t) return false;
    }
    return true;
  });

  const tbody = document.getElementById("rdTbody");
  const empty = document.getElementById("rdEmpty");
  const tbl = document.getElementById("rdTable");
  document.getElementById("rdCount").innerHTML =
    `<strong>${filtered.length}</strong> of ${_rdState.rows.length} successors`;

  if (!filtered.length) {
    tbody.innerHTML = "";
    empty.style.display = "block";
    tbl.style.display = "none";
    return;
  }
  empty.style.display = "none";
  tbl.style.display = "";

  tbody.innerHTML = filtered.map(r => {
    const pctCls = r.overallPct === null ? "lo" : r.overallPct >= 100 ? "done" : r.overallPct >= 75 ? "hi" : r.overallPct >= 50 ? "mid" : "lo";
    const pctVal = r.overallPct === null ? "—" : r.overallPct + "%";
    const pctWidth = r.overallPct === null ? 0 : Math.min(100, r.overallPct);
    const statusCls = r.learningStatus === "Completed" ? "completed" : r.learningStatus === "In Progress" ? "inprogress" : "notstarted";
    const courseHtml = r.courses.length
      ? r.courses.map(c => `<div class="rd-course-line"><span>${_esc(c.name)}</span><span style="flex-shrink:0;">${c.pct!==null?c.pct+'%':'—'}</span></div>`).join("")
      : '<span style="color:var(--t3);font-size:11px;">No learning data</span>';
    return `<tr>
      <td><div class="rd-emp"><span class="rd-emp-name">${_esc(r.succName)}</span><span class="rd-emp-pos">${_esc(r.successorPosition)}</span></div></td>
      <td style="font-family:monospace;font-size:11px;color:var(--t2);">${_esc(r.succNik || "—")}</td>
      <td>${_esc(r.successorPosition)}</td>
      <td>${_esc(r.successorForPos)}${r.isCLevel?' <span class="rd-badge clevel" style="margin-left:4px;">C</span>':''}</td>
      <td>${_esc(r.department)}</td>
      <td><span class="rd-badge ${r.readinessCls}">${_esc(r.readiness)}</span></td>
      <td><div class="rd-courses">${courseHtml}</div></td>
      <td><div class="rd-progress"><div class="rd-progress-track"><div class="rd-progress-fill ${pctCls}" style="width:${pctWidth}%;"></div></div><span class="rd-progress-val">${pctVal}</span></div></td>
      <td><span class="rd-badge ${statusCls}">${_esc(r.learningStatus)}</span></td>
      <td style="font-size:11px;color:var(--t3);white-space:nowrap;">${_esc(r.lastUpdated || "—")}</td>
    </tr>`;
  }).join("");
}

// Export current filtered view to Excel
function exportReadinessDetail(){
  const q = (document.getElementById("rdSearch").value || "").toLowerCase().trim();
  const dept = document.getElementById("rdDept").value;
  const level = document.getElementById("rdLevel").value;
  const learnStatus = document.getElementById("rdLearnStatus").value;
  const pctThresh = document.getElementById("rdPct").value;
  const filtered = _rdState.rows.filter(r => {
    if (q && !(r.succName.toLowerCase().includes(q) || (r.succNik||"").toLowerCase().includes(q))) return false;
    if (dept && r.department !== dept) return false;
    if (level === "c" && !r.isCLevel) return false;
    if (level === "non" && r.isCLevel) return false;
    if (learnStatus && r.learningStatus !== learnStatus) return false;
    if (pctThresh !== "") {
      const t = parseFloat(pctThresh);
      if (r.overallPct === null) return t === 0;
      if (t === 100) return r.overallPct >= 100;
      if (t === 0) return r.overallPct === 0;
      if (r.overallPct < t) return false;
    }
    return true;
  });
  const exportRows = filtered.map(r => ({
    "Employee Name": r.succName,
    "Employee ID": r.succNik || "",
    "Current Position": r.successorPosition,
    "Successor For": r.successorForPos,
    "Department": r.department,
    "Business Unit": r.businessUnit,
    "C-Level": r.isCLevel ? "Yes" : "No",
    "Readiness": r.readiness,
    "Courses Count": r.courses.length,
    "Course Names": r.courses.map(c=>c.name).join("; "),
    "Overall Completion %": r.overallPct === null ? "" : r.overallPct,
    "Learning Status": r.learningStatus,
    "Last Updated": r.lastUpdated || ""
  }));
  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, _rdState.category);
  const fname = `${_rdState.category.replace(/\s+/g,"_")}_successors_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fname);
}

// Close on Escape
document.addEventListener("keydown", e => {
  if (e.key === "Escape" && document.getElementById("rdBackdrop").classList.contains("show")) closeReadinessDetail();
  if (e.key === "Escape" && document.getElementById("dualBackdrop").classList.contains("show")) closeMyLearningDualUpload();
});

// ═══════════════════════════════════════════════════════════════════════════
// FEATURE 2: MYLEARNING DUAL-FILE UPLOAD
// ═══════════════════════════════════════════════════════════════════════════
// Two slots (Transcript + Course Completion) merged by NIK → Email → Name,
// then fed into the existing MyLearning ingest pipeline.
// ═══════════════════════════════════════════════════════════════════════════
const _dualState = { transcript: null, progress: null };

function openMyLearningDualUpload(){
  document.getElementById("dualBackdrop").classList.add("show");
  document.body.style.overflow = "hidden";
  _resetDualState();
}
function closeMyLearningDualUpload(){
  document.getElementById("dualBackdrop").classList.remove("show");
  document.body.style.overflow = "";
}
function _resetDualState(){
  _dualState.transcript = null; _dualState.progress = null;
  ["transcript","progress"].forEach(slot => {
    document.getElementById("dualSlot_"+slot).classList.remove("has-file");
    document.getElementById("dualMeta_"+slot).innerHTML = "";
  });
  document.getElementById("dualProcessBtn").disabled = true;
  document.getElementById("dualResult").innerHTML = "";
  document.getElementById("dualProgress").style.display = "none";
}

function pickDualFile(slot){
  const inp = document.getElementById("dualFileInput");
  inp.onchange = e => {
    if (e.target.files[0]) _loadDualFile(slot, e.target.files[0]);
    inp.value = ""; inp.onchange = null;
  };
  inp.click();
}

function _loadDualFile(slot, file){
  if (!/\.(xlsx|xls|xlsm|csv|tsv)$/i.test(file.name)) {
    _showDualResult("error", `<strong>${file.name}</strong> isn't a supported file type. Use .xlsx, .xls, or .csv.`);
    return;
  }
  const r = new FileReader();
  r.onload = ev => {
    try {
      const wb = XLSX.read(ev.target.result, {type:"array", raw:false});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {defval:"", raw:false});
      if (!rows.length) { _showDualResult("error", `<strong>${file.name}</strong> contains no data rows.`); return; }
      _dualState[slot] = { file, rows, headers: Object.keys(rows[0]) };
      const card = document.getElementById("dualSlot_"+slot);
      card.classList.add("has-file");
      document.getElementById("dualMeta_"+slot).innerHTML =
        `<strong>${_esc(file.name)}</strong> • ${rows.length.toLocaleString()} rows · ${Object.keys(rows[0]).length} columns`;
      // Enable process if both loaded
      const both = _dualState.transcript && _dualState.progress;
      document.getElementById("dualProcessBtn").disabled = !both;
      if (both) {
        _showDualResult("success", `Both files loaded. Click <strong>Process &amp; Load</strong> to merge and validate.`);
      } else {
        _showDualResult("warn", `File ${slot==='transcript'?'1':'2'} loaded. Add the other file to enable processing.`);
      }
    } catch (err) {
      _showDualResult("error", `Could not read <strong>${file.name}</strong>: ${err.message}`);
    }
  };
  r.readAsArrayBuffer(file);
}

function _showDualResult(level, msg){
  document.getElementById("dualResult").innerHTML =
    `<div class="dual-result-card ${level}">${msg}</div>`;
}

// Find which canonical field each header maps to (via FIELD_REGISTRY aliases)
function _detectKeyHeaders(headers){
  const keys = { nik: null, email: null, name: null, course: null };
  headers.forEach(h => {
    const m = _fuzzyMatch(h);
    if (!m) return;
    if (m.canonical === "NIK"            && !keys.nik)    keys.nik = h;
    if (m.canonical === "Email"          && !keys.email)  keys.email = h;
    if (m.canonical === "Employee Name"  && !keys.name)   keys.name = h;
    if (m.canonical === "Training Name"  && !keys.course) keys.course = h;
  });
  return keys;
}

// Step-driven progress update
function _setDualProgress(label, pct){
  const wrap = document.getElementById("dualProgress");
  wrap.style.display = "block";
  document.getElementById("dualProgressLbl").textContent = label;
  document.getElementById("dualProgressFill").style.width = pct + "%";
}

// ── The actual merge: join Transcript + Progress on NIK→Email→Name ──────
async function processDualUpload(){
  if (!_dualState.transcript || !_dualState.progress) return;
  const btn = document.getElementById("dualProcessBtn");
  btn.disabled = true;

  // Use a microtask + setTimeout so the progress bar is visibly animated
  _setDualProgress("Validating file structure…", 15);
  await new Promise(r => setTimeout(r, 80));

  const T = _dualState.transcript, P = _dualState.progress;
  const tKeys = _detectKeyHeaders(T.headers);
  const pKeys = _detectKeyHeaders(P.headers);

  // Validate: both files must have at least one identity column
  if (!tKeys.nik && !tKeys.email && !tKeys.name) {
    _setDualProgress("", 0); document.getElementById("dualProgress").style.display = "none";
    _showDualResult("error", `Transcript file has no recognizable identity column (NIK, Email, or Employee Name).<br>Detected headers: <code>${T.headers.slice(0,8).join(", ")}…</code>`);
    btn.disabled = false; return;
  }
  if (!pKeys.nik && !pKeys.email && !pKeys.name) {
    _setDualProgress("", 0); document.getElementById("dualProgress").style.display = "none";
    _showDualResult("error", `Progress file has no recognizable identity column (NIK, Email, or Employee Name).`);
    btn.disabled = false; return;
  }

  _setDualProgress("Indexing transcript file by employee…", 35);
  await new Promise(r => setTimeout(r, 80));

  // Build indices on the Progress file (smaller-or-equal expected) keyed by NIK/Email/Name
  // Multiple progress rows per employee allowed (one per course).
  const progressByNik = {}, progressByEmail = {}, progressByName = {};
  P.rows.forEach((row, i) => {
    const nik = pKeys.nik   ? _norm(row[pKeys.nik])   : "";
    const eml = pKeys.email ? _norm(row[pKeys.email]) : "";
    const nm  = pKeys.name  ? _norm(row[pKeys.name])  : "";
    if (nik) (progressByNik[nik]   = progressByNik[nik]   || []).push(row);
    if (eml) (progressByEmail[eml] = progressByEmail[eml] || []).push(row);
    if (nm)  (progressByName[nm]   = progressByName[nm]   || []).push(row);
  });

  _setDualProgress("Merging records by NIK → Email → Name…", 60);
  await new Promise(r => setTimeout(r, 100));

  // Track stats
  const stats = {
    transcriptRows: T.rows.length,
    progressRows: P.rows.length,
    mergedRows: 0,
    matchedByNik: 0,
    matchedByEmail: 0,
    matchedByName: 0,
    transcriptOrphans: 0,    // transcript rows with no progress match
    progressOrphansCount: 0, // progress rows that never matched any transcript
    duplicatesInTranscript: 0,
    duplicatesInProgress: 0,
  };

  // Detect duplicates within each file (same employee + same course twice)
  const tKey = r => {
    const id = (tKeys.nik && r[tKeys.nik]) || (tKeys.email && r[tKeys.email]) || (tKeys.name && r[tKeys.name]) || "";
    const c  = tKeys.course ? r[tKeys.course] : "";
    return _norm(id) + "|" + _norm(c);
  };
  const pKey = r => {
    const id = (pKeys.nik && r[pKeys.nik]) || (pKeys.email && r[pKeys.email]) || (pKeys.name && r[pKeys.name]) || "";
    const c  = pKeys.course ? r[pKeys.course] : "";
    return _norm(id) + "|" + _norm(c);
  };
  const tSeen = {}, pSeen = {};
  T.rows.forEach(r => { const k = tKey(r); if (k.length>1) { if (tSeen[k]) stats.duplicatesInTranscript++; else tSeen[k] = true; } });
  P.rows.forEach(r => { const k = pKey(r); if (k.length>1) { if (pSeen[k]) stats.duplicatesInProgress++; else pSeen[k] = true; } });

  // For each transcript row, find best progress match.
  // Preferred join: (employee × course) — same course in both files merges into one row.
  // If progress has multiple rows for same employee but no matching course, all become extra rows.
  const merged = [];
  const matchedProgressKeys = new Set();
  T.rows.forEach(trow => {
    const nik = tKeys.nik   ? _norm(trow[tKeys.nik])   : "";
    const eml = tKeys.email ? _norm(trow[tKeys.email]) : "";
    const nm  = tKeys.name  ? _norm(trow[tKeys.name])  : "";
    const courseN = tKeys.course ? _norm(trow[tKeys.course]) : "";

    // Resolve employee in progress index
    let candidates = null, joinReason = null;
    if (nik && progressByNik[nik])       { candidates = progressByNik[nik];   joinReason = "nik"; }
    else if (eml && progressByEmail[eml]){ candidates = progressByEmail[eml]; joinReason = "email"; }
    else if (nm && progressByName[nm])   { candidates = progressByName[nm];   joinReason = "name"; }

    if (!candidates) {
      // Transcript orphan — keep with empty progress fields
      merged.push({ ...trow, _mergeSource: "transcript_only" });
      stats.transcriptOrphans++;
      return;
    }

    // Prefer course-matched candidate; else fall back to first candidate
    let prow = null;
    if (courseN) {
      prow = candidates.find(pr => _norm(pr[pKeys.course] || "") === courseN);
    }
    if (!prow) prow = candidates[0];

    // Mark this progress row as used
    const pk = pKey(prow);
    matchedProgressKeys.add(pk);

    // Merge: transcript values win unless empty, then progress fills
    const combined = { ...trow };
    Object.keys(prow).forEach(k => {
      const tv = combined[k];
      const pv = prow[k];
      if ((tv === undefined || tv === null || tv === "" || tv === "-") && pv !== "" && pv !== null && pv !== undefined) {
        combined[k] = pv;
      } else if (!(k in combined)) {
        combined[k] = pv;
      }
    });
    combined._mergeSource = "both";
    merged.push(combined);
    if (joinReason === "nik")   stats.matchedByNik++;
    if (joinReason === "email") stats.matchedByEmail++;
    if (joinReason === "name")  stats.matchedByName++;
  });

  // Progress orphans: rows that never matched any transcript row
  P.rows.forEach(prow => {
    const pk = pKey(prow);
    if (!matchedProgressKeys.has(pk)) {
      merged.push({ ...prow, _mergeSource: "progress_only" });
      stats.progressOrphansCount++;
    }
  });
  stats.mergedRows = merged.length;

  _setDualProgress("Loading into MyLearning source…", 85);
  await new Promise(r => setTimeout(r, 100));

  // Hand off to the existing ingest pipeline.
  // We need to give the mapper a representative row so headers can be auto-detected.
  // Use the union of all keys across merged rows so the mapper sees every column.
  const allHeaders = new Set();
  merged.forEach(r => Object.keys(r).forEach(k => { if (k !== "_mergeSource") allHeaders.add(k); }));
  const headersArr = Array.from(allHeaders);
  // Normalize: ensure every row has every header (so XLSX.utils.json_to_sheet doesn't drop columns)
  const normalized = merged.map(r => {
    const out = {};
    headersArr.forEach(h => out[h] = r[h] !== undefined ? r[h] : "");
    return out;
  });

  // Auto-detect mapping
  const autoMapping = {};
  const savedOverrides = window.MAPPING_CONFIG.loadOverrides
    ? window.MAPPING_CONFIG.loadOverrides("mylearning") : null;
  headersArr.forEach(h => {
    if (savedOverrides && Object.prototype.hasOwnProperty.call(savedOverrides, h)) {
      const canon = savedOverrides[h];
      autoMapping[h] = { canonical: canon || null, confidence: canon ? "saved" : "ignored" };
      return;
    }
    const match = _fuzzyMatch(h);
    autoMapping[h] = match ? { canonical: match.canonical, confidence: match.confidence } : { canonical: null, confidence: "unmapped" };
  });

  // Synthesize a fake File for display
  const syntheticName = `merged_${T.file.name.replace(/\.[^.]+$/,"")}__${P.file.name.replace(/\.[^.]+$/,"")}.xlsx`;
  const ws = XLSX.utils.json_to_sheet(normalized);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "MyLearning");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const syntheticFile = new File([buf], syntheticName, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  _pendingSource = { sourceId: "mylearning", file: syntheticFile, raw: normalized, headers: headersArr, autoMapping };

  _setDualProgress("Done", 100);
  await new Promise(r => setTimeout(r, 200));

  // Show summary
  const matchedTotal = stats.matchedByNik + stats.matchedByEmail + stats.matchedByName;
  const warnings = [];
  if (stats.duplicatesInTranscript) warnings.push(`${stats.duplicatesInTranscript} duplicate row${stats.duplicatesInTranscript>1?"s":""} in transcript`);
  if (stats.duplicatesInProgress)   warnings.push(`${stats.duplicatesInProgress} duplicate row${stats.duplicatesInProgress>1?"s":""} in progress file`);
  if (stats.transcriptOrphans)      warnings.push(`${stats.transcriptOrphans} transcript row${stats.transcriptOrphans>1?"s":""} with no progress match`);
  if (stats.progressOrphansCount)   warnings.push(`${stats.progressOrphansCount} progress row${stats.progressOrphansCount>1?"s":""} with no transcript match`);

  const level = warnings.length ? "warn" : "success";
  _showDualResult(level, `
    <strong>Merge complete.</strong> ${stats.mergedRows.toLocaleString()} unified rows produced
    from ${stats.transcriptRows.toLocaleString()} transcript + ${stats.progressRows.toLocaleString()} progress rows.<br>
    <div style="margin-top:6px;font-size:11px;">
      Matched ${matchedTotal.toLocaleString()} pairs
      ${stats.matchedByNik?`<span style="color:#065F46;">• ${stats.matchedByNik} by NIK</span>`:""}
      ${stats.matchedByEmail?`<span style="color:#065F46;">• ${stats.matchedByEmail} by Email</span>`:""}
      ${stats.matchedByName?`<span style="color:#92400E;">• ${stats.matchedByName} by Name (fallback)</span>`:""}.
    </div>
    ${warnings.length ? `<div style="margin-top:6px;font-size:11px;">⚠ ${warnings.join("; ")}.</div>` : ""}
    <div style="margin-top:8px;">Opening column mapper next…</div>
  `);

  await new Promise(r => setTimeout(r, 600));
  closeMyLearningDualUpload();
  // Hand off to existing mapper UI
  openMapper(_pendingSource, /*isNew=*/true);
}

// setupDualDragDrop moved to useEffect in App.jsx


// ── Expose all public functions to window for dynamic innerHTML onclick handlers ──
const _publicFns = {
  downloadTemplate, renderHub, pickFileFor, ingestFile, openMapperFor, clearSource,
  openMapper, closeMapper, confirmMapping, _updateMapBadge, validateSource,
  openIssuesPanel, closeIssuesPanel, _setIssuesTab, buildAndEnter, resetApp,
  applyFilters, switchTab, openSnapModal, closeSnapModal, setSnapFilter,
  setSnapSuccFilter, filterSnapModal, renderPositions, setPF, openPosDetail,
  closePosDetail, openSuccPopup, closeSuccPopup, runSimulation, _filterSimPositions,
  _toggleSimDetail, detailSearchHandler, renderDetail, applyOverviewUniqFilter,
  openReadinessDetail, closeReadinessDetail, renderReadinessDetailRows,
  exportReadinessDetail, openMyLearningDualUpload, closeMyLearningDualUpload,
  pickDualFile, _loadDualFile, processDualUpload, scPreview, scGenerate,
  _pgGo, _ghFilter, esPage, closeModal, confirmLoad,
};
Object.assign(window, _publicFns);

// ─────────────────────────────────────────────────────────────────────────────
// AUTO-LOAD: fetch a file from the public directory and build the dashboard
// Call this instead of renderHub() to skip the upload screen entirely.
// ─────────────────────────────────────────────────────────────────────────────
async function autoLoadFromUrl(url, sourceId = 'sap') {
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', raw: false });
    const idx = wb.SheetNames.findIndex(
      s => s.toLowerCase() !== 'instructions' && s.toLowerCase() !== 'instruction'
    );
    const ws = wb.Sheets[wb.SheetNames[idx >= 0 ? idx : 0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
    if (!raw.length) throw new Error('No data rows in file');

    const headers = Object.keys(raw[0]);
    const savedOverrides = window.MAPPING_CONFIG && window.MAPPING_CONFIG.loadOverrides
      ? window.MAPPING_CONFIG.loadOverrides(sourceId) : null;

    const mapping = {};
    headers.forEach(h => {
      if (savedOverrides && Object.prototype.hasOwnProperty.call(savedOverrides, h)) {
        const canon = savedOverrides[h];
        if (canon) mapping[h] = canon;
      } else {
        const match = _fuzzyMatch(h);
        if (match) mapping[h] = match.canonical;
      }
    });

    const s = SOURCES[sourceId];
    s.file = { name: url.split('/').pop() };
    s.raw = raw;
    s.headers = headers;
    s.mapping = mapping;
    validateSource(sourceId);

    ALL = buildIntegratedData();
    if (!ALL.length) throw new Error('No rows after integration');

    const topFile = document.getElementById('topFile');
    const topRec  = document.getElementById('topRec');
    if (topFile) topFile.textContent = s.file.name;
    if (topRec)  topRec.textContent  = ALL.length + ' rows';

    document.getElementById('upload-screen').style.display = 'none';
    document.getElementById('dashboard').style.display     = 'block';

    buildFilters();
    FILTERED = [...ALL];
    renderAll();
  } catch (err) {
    console.warn('Auto-load failed (' + url + '):', err.message);
    // Fall back to the manual upload screen
    renderHub();
  }
}
window.autoLoadFromUrl = autoLoadFromUrl;
