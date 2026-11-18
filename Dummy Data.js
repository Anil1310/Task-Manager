// ============================================================
// DUMMY DATA — Seed sample users, projects, and tasks
// ============================================================
function seedDefaultData() {
  const ss = getSpreadsheet();

  // --- Seed Users ---
  var userSheet = getOrCreateSheet(ss, 'Users', [
    'Username', 'Password', 'Name', 'Role', 'Initials', 'Color', 'Department', 'Manager', 'Project', 'Email', 'Mobile', 'Photo'
  ]);
  var usersData = userSheet.getDataRange().getValues();
  var hasOldUsers = usersData.some(function(row) {
    var username = String(row[0]).toLowerCase().trim();
    return username === 'rajesh' || username === 'deepa' || username === 'priya';
  });

  var users = [
    ['admin', 'pass123', 'Admin', 'Admin', 'SA', '#3c6ed3', 'IT', 'Rajeev', 'Skyline Towers', 'automationbyanil@gmail.com', '', ''],
    ['rajeev', 'pass123', 'Rajeev Channa', 'MD', 'RC', '#06db9b', 'MDO', 'Rajeev', 'Head  Office', '', '', ''],
    ['anil', 'pass123', 'Anil Chahar', 'Manager', 'AC', '#0d9488', 'IT', 'Rajeev', 'MD', '', '8448189027', ''],
    ['shabnam', 'pass123', 'Sbabanm Das', 'Manager', 'SD', '#871717', 'HR', 'Rajeev', 'HR', 'shabnam@aarcon.co', '9266344129', ''],
    ['varsha', 'pass123', 'Varsha', 'User', 'AK', '#d97706', 'IT', 'Rajeev', 'HO', 'varsha.aarcon932026@gmail.com', '9910269491', ''],
    ['gopal', 'pass123', 'Gopal', 'User', 'G', '#0d9488', 'Account', 'Rajeev', 'HO', 'gopal@gmail.com', '9910269492', ''],
    ['nisha', 'pass123', 'NIsha Yadav', 'User', 'NY', '#06db9b', 'Digital Marketing', 'Rajeev', 'HO', 'nisha@gmail.com', '9910269493', ''],
    ['ahmad', 'pass123', 'Ahmad', 'Manager', 'A', '#871717', 'Purchase', 'Rajeev', 'HO', 'ahmad@gmail.com', '9910269494', ''],
    ['jitender', 'pass123', 'Jinteder', 'User', 'J', '#d97706', 'Purchase', 'Ahmad', 'HO', 'jitender@gmail.com', '9910269495', ''],
    ['sahil', 'pass123', 'Sahil', 'User', 'S', '#d97706', 'Purchase', 'Ahmad', 'HO', 'sahil@gmail.com', '9910269496', ''],
    ['anus', 'pass123', 'Anus', 'User', 'A', '#d97706', 'Purchase', 'Ahmad', 'HO', 'anus@gmail.com', '9910269497', '']
  ];

  if (userSheet.getLastRow() < 2 || hasOldUsers) {
    if (userSheet.getLastRow() >= 2) {
      userSheet.deleteRows(2, userSheet.getLastRow() - 1);
    }
    users.forEach(function(row) { userSheet.appendRow(row); });
  }

  // --- Seed Projects ---
  var projectSheet = getOrCreateSheet(ss, 'Projects', [
    'Name', 'Location', 'Progress', 'Budget'
  ]);
  if (projectSheet.getLastRow() < 2) {
    var projects = [
      ['Skyline Towers', 'Downtown West', 68, '12.5 Cr'],
      ['Green Valley Residency', 'Sector 21, North', 42, '8.2 Cr'],
      ['Metro Bridge Overpass', 'River Zone', 85, '24.0 Cr'],
      ['Sunrise Mall', 'Highway East', 23, '18.7 Cr']
    ];
    projects.forEach(function(row) { projectSheet.appendRow(row); });
  }

  // --- Seed Tasks ---
  var taskSheet = getOrCreateSheet(ss, 'Tasks', [
    'Id', 'Title', 'Desc', 'Project', 'Assignee', 'Email', 'Assigned By', 'Priority',
    'Status', 'Progress', 'Deadline', 'Created', 'Comments'
  ]);
  var tasksData = taskSheet.getDataRange().getValues();
  var hasOldTasks = tasksData.some(function(row) {
    var assignee = String(row[4]).toLowerCase().trim();
    var emailVal = String(row[5]).toLowerCase().trim();
    return assignee === 'deepa' || assignee === 'vikram' || assignee === 'priya' ||
           emailVal === 'rajeev' || emailVal === 'admin' || emailVal === 'anil' || emailVal === 'shabnam';
  });

  // Force re-seed if the Tasks sheet is empty, has shifted/old tasks, or is incomplete (fewer than 100 rows)
  if (taskSheet.getLastRow() < 100 || hasOldTasks) {
    if (taskSheet.getLastRow() >= 2) {
      taskSheet.deleteRows(2, taskSheet.getLastRow() - 1);
    }
    var tasks = [
      [1, 'Foundation waterproofing — Block A', 'Apply double-coat waterproofing membrane on all foundation walls.', 'Skyline Towers', 'anil', 'rajeev', 'High', 'In Progress', 65, '2026-08-10', '2026-05-01', '["{\\\"username\\\":\\\"anil\\\",\\\"author\\\":\\\"Anil Chahar\\\",\\\"role\\\":\\\"Manager\\\",\\\"text\\\":\\\"Material delivered\\\",\\\"timestamp\\\":\\\"01 May, 10:30 am\\\"}\", \"{\\\"username\\\":\\\"anil\\\",\\\"author\\\":\\\"Anil Chahar\\\",\\\"role\\\":\\\"Manager\\\",\\\"text\\\":\\\"First coat completed\\\",\\\"timestamp\\\":\\\"05 May, 04:15 pm\\\"}\"]'],
      [2, 'Steel reinforcement inspection — Level 3', 'Inspect all rebar placement on Level 3 before concrete pour.', 'Skyline Towers', 'varsha', 'anil', 'High', 'Pending', 0, '2026-07-15', '2026-05-03', '[]'],
      [3, 'Concrete pour scheduling — Basement 2', 'Plan and schedule the concrete pour for Basement 2 slab.', 'Green Valley Residency', 'nisha', 'rajeev', 'Medium', 'In Progress', 40, '2026-09-22', '2026-05-02', '["{\\\"username\\\":\\\"nisha\\\",\\\"author\\\":\\\"NIsha Yadav\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Mix design approved\\\",\\\"timestamp\\\":\\\"03 May, 11:30 am\\\"}\"]'],
      [4, 'Safety audit — Crane operations', 'Complete full safety audit of tower crane operations.', 'Metro Bridge Overpass', 'gopal', 'shabnam', 'High', 'Completed', 100, '2026-06-10', '2026-05-28', '["{\\\"username\\\":\\\"gopal\\\",\\\"author\\\":\\\"Gopal\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Audit completed\\\",\\\"timestamp\\\":\\\"28 May, 02:40 pm\\\"}\"]'],
      [5, 'MEP coordination meeting', 'Arrange coordination meeting between MEP subcontractors.', 'Sunrise Mall', 'jitender', 'ahmad', 'Medium', 'Under Review', 80, '2026-07-18', '2026-05-01', '["{\\\"username\\\":\\\"jitender\\\",\\\"author\\\":\\\"Jinteder\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Meeting held\\\",\\\"timestamp\\\":\\\"02 May, 09:00 am\\\"}\"]'],
      [6, 'Site survey and layout marking — Phase 2', 'Complete total station survey for Phase 2 boundary.', 'Green Valley Residency', 'sahil', 'ahmad', 'Medium', 'Pending', 0, '2026-10-25', '2026-05-05', '[]'],
      [7, 'Quality check — curing compound application', 'Verify proper application of curing compound.', 'Skyline Towers', 'anus', 'ahmad', 'Low', 'In Progress', 50, '2026-08-28', '2026-05-04', '[]'],
      [8, 'Bridge girder launching — Span 3', 'Supervise launching of precast girders.', 'Metro Bridge Overpass', 'nisha', 'rajeev', 'High', 'In Progress', 30, '2026-07-16', '2026-05-30', '["{\\\"username\\\":\\\"nisha\\\",\\\"author\\\":\\\"NIsha Yadav\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Girders transported\\\",\\\"timestamp\\\":\\\"30 May, 02:40 pm\\\"}\"]'],
      [9, 'Material procurement — plumbing fixtures', 'Finalize procurement list for plumbing fixtures.', 'Sunrise Mall', 'anil', 'rajeev', 'Low', 'Pending', 0, '2026-09-05', '2026-07-06', '[]'],
      [10, 'Piling work inspection — Zone C', 'Inspect bored piling work in Zone C.', 'Green Valley Residency', 'varsha', 'anil', 'High', 'Pending', 0, '2026-08-08', '2026-06-25', '["{\\\"username\\\":\\\"varsha\\\",\\\"author\\\":\\\"Varsha\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Delay due to hard rock\\\",\\\"timestamp\\\":\\\"26 Jun, 10:00 am\\\"}\"]'],
      [11, 'Electrical panel installation — Floor 1-4', 'Install main distribution panels.', 'Skyline Towers', 'gopal', 'shabnam', 'Medium', 'Completed', 100, '2026-06-09', '2026-06-20', '["{\\\"username\\\":\\\"gopal\\\",\\\"author\\\":\\\"Gopal\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Panels Floor 1-2 installed\\\",\\\"timestamp\\\":\\\"20 Jun, 03:20 pm\\\"}\"]'],
      [12, 'Landscaping review — Entrance plaza', 'Review landscaping plan for entrance plaza.', 'Sunrise Mall', 'anus', 'ahmad', 'Low', 'Completed', 100, '2026-06-12', '2026-07-01', '["{\\\"username\\\":\\\"anus\\\",\\\"author\\\":\\\"Anus\\\",\\\"role\\\":\\\"User\\\",\\\"text\\\":\\\"Plan approved\\\",\\\"timestamp\\\":\\\"02 Jul, 11:00 am\\\"}\"]'],
      [13, 'Cement batch delivery verification', 'Inspect material quality certificates and record batch numbers on arrival.', 'Skyline Towers', 'admin', 'rajeev', 'Low', 'Under Review', 87, '2026-06-06', '2026-05-09', '["{\\\"username\\\": \\\"admin\\\", \\\"author\\\": \\\"Admin\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"18 may, 12:00 am\\\"}"]'],
      [14, 'Boundary wall layout marking — Zone 7', 'Set out coordinates and install marking stakes for the boundary wall.', 'Skyline Towers', 'admin', 'rajeev', 'High', 'Pending', 0, '2026-06-11', '2026-05-08', '[]'],
      [15, 'Water supply header installation — Zone 1', 'Inspect PPR pipe welding and pressure testing of header lines.', 'Green Valley Residency', 'admin', 'rajeev', 'Low', 'Completed', 100, '2026-06-26', '2026-05-09', '["{\\\"username\\\": \\\"admin\\\", \\\"author\\\": \\\"Admin\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"10 may, 12:00 am\\\"}"]'],
      [16, 'Waste debris clearance coordination', 'Organize waste collection vehicles and record disposal manifesting.', 'Sunrise Mall', 'admin', 'rajeev', 'Medium', 'Under Review', 84, '2026-06-18', '2026-05-08', '[]'],
      [17, 'Basement flooring screed curing — Zone 2', 'Supervise floor hardener application and spray curing of the screed slab.', 'Metro Bridge Overpass', 'admin', 'rajeev', 'Medium', 'Under Review', 81, '2026-07-13', '2026-05-25', '["{\\\"username\\\": \\\"admin\\\", \\\"author\\\": \\\"Admin\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"01 jun, 12:00 am\\\"}"]'],
      [18, 'HVAC duct installation — Level 5', 'Verify hanger spacing, sleeve layout, and damper positions for ducts.', 'Metro Bridge Overpass', 'admin', 'rajeev', 'Low', 'In Progress', 18, '2026-06-06', '2026-05-03', '["{\\\"username\\\": \\\"admin\\\", \\\"author\\\": \\\"Admin\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"05 may, 12:00 am\\\"}"]'],
      [19, 'Concrete compressive strength test — Batch 2', 'Perform slump test and prepare concrete cylinder molds for laboratory testing.', 'Sunrise Mall', 'admin', 'rajeev', 'Medium', 'Completed', 100, '2026-07-04', '2026-05-22', '[]'],
      [20, 'Exterior glass facade installation — Block 4', 'Coordinate structural glazing brackets installation and safety netting.', 'Metro Bridge Overpass', 'admin', 'rajeev', 'Low', 'Pending', 0, '2026-07-20', '2026-05-21', '[]'],
      [21, 'Rebar cage inspection — Column B9', 'Verify steel reinforcement diameters and spacing for basement Column B9.', 'Green Valley Residency', 'admin', 'rajeev', 'High', 'Completed', 100, '2026-06-20', '2026-05-14', '["{\\\"username\\\": \\\"admin\\\", \\\"author\\\": \\\"Admin\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"23 may, 12:00 am\\\"}"]'],
      [22, 'Fine sand sieve analysis', 'Perform silt content and sieve analysis tests on incoming sand stock.', 'Skyline Towers', 'admin', 'rajeev', 'High', 'Pending', 0, '2026-07-06', '2026-05-27', '[]'],
      [23, 'Basement flooring screed curing — Zone 5', 'Supervise floor hardener application and spray curing of the screed slab.', 'Skyline Towers', 'rajeev', 'admin', 'High', 'Under Review', 86, '2026-07-12', '2026-05-22', '[]'],
      [24, 'Cement batch delivery verification', 'Inspect material quality certificates and record batch numbers on arrival.', 'Green Valley Residency', 'rajeev', 'admin', 'Medium', 'In Progress', 41, '2026-07-19', '2026-05-25', '["{\\\"username\\\": \\\"rajeev\\\", \\\"author\\\": \\\"Rajeev\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"04 jun, 12:00 am\\\"}"]'],
      [25, 'Drywall framing inspection — Level 7', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Metro Bridge Overpass', 'rajeev', 'admin', 'High', 'In Progress', 75, '2026-06-11', '2026-05-17', '["{\\\"username\\\": \\\"rajeev\\\", \\\"author\\\": \\\"Rajeev\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"19 may, 12:00 am\\\"}"]'],
      [26, 'Drainage line layout design — Zone 3', 'Review slope specifications and finalize storm water line routes.', 'Sunrise Mall', 'rajeev', 'admin', 'Low', 'Pending', 0, '2026-06-27', '2026-05-14', '[]'],
      [27, 'Main panel board terminal check — Block 8', 'Check torque on line wire connections and ground resistance values.', 'Metro Bridge Overpass', 'rajeev', 'admin', 'Low', 'Pending', 0, '2026-06-19', '2026-05-23', '[]'],
      [28, 'Fine sand sieve analysis', 'Perform silt content and sieve analysis tests on incoming sand stock.', 'Metro Bridge Overpass', 'rajeev', 'admin', 'Low', 'Under Review', 83, '2026-06-27', '2026-05-11', '[]'],
      [29, 'Geotechnical soil testing — Phase 5', 'Perform borehole drilling and soil compaction tests for Phase 5 foundations.', 'Green Valley Residency', 'rajeev', 'admin', 'Low', 'Pending', 0, '2026-07-28', '2026-05-29', '[]'],
      [30, 'Precast beam launching — Span 9', 'Supervise crane positioning and rigging for precast girder alignment.', 'Green Valley Residency', 'rajeev', 'admin', 'High', 'Under Review', 85, '2026-07-11', '2026-05-19', '["{\\\"username\\\": \\\"rajeev\\\", \\\"author\\\": \\\"Rajeev\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"29 may, 12:00 am\\\"}"]'],
      [31, 'Internal plastering audit — Floor 8', 'Inspect surface smoothness and check plaster thickness for internal walls.', 'Skyline Towers', 'rajeev', 'admin', 'High', 'Under Review', 89, '2026-06-01', '2026-05-09', '[]'],
      [32, 'Fire hydrant riser testing — Phase 2', 'Conduct hydrostatic pressure testing on wet riser pipe lines.', 'Skyline Towers', 'rajeev', 'admin', 'Low', 'Completed', 100, '2026-06-21', '2026-05-28', '["{\\\"username\\\": \\\"rajeev\\\", \\\"author\\\": \\\"Rajeev\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"06 jun, 12:00 am\\\"}"]'],
      [33, 'Waste debris clearance coordination', 'Organize waste collection vehicles and record disposal manifesting.', 'Green Valley Residency', 'anil', 'rajeev', 'Low', 'Completed', 100, '2026-06-18', '2026-05-19', '[]'],
      [34, 'Main panel board terminal check — Block 7', 'Check torque on line wire connections and ground resistance values.', 'Green Valley Residency', 'anil', 'rajeev', 'Low', 'In Progress', 49, '2026-06-26', '2026-05-14', '["{\\\"username\\\": \\\"anil\\\", \\\"author\\\": \\\"Anil\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"23 may, 12:00 am\\\"}"]'],
      [35, 'Tile layout alignment check — Lobby 2', 'Review tile spacing, layout consistency, and grout line widths.', 'Green Valley Residency', 'anil', 'rajeev', 'High', 'Pending', 0, '2026-06-02', '2026-05-12', '[]'],
      [36, 'Fire hydrant riser testing — Phase 9', 'Conduct hydrostatic pressure testing on wet riser pipe lines.', 'Green Valley Residency', 'anil', 'rajeev', 'Low', 'In Progress', 10, '2026-07-03', '2026-05-04', '[]'],
      [37, 'Boundary wall layout marking — Zone 1', 'Set out coordinates and install marking stakes for the boundary wall.', 'Metro Bridge Overpass', 'anil', 'rajeev', 'High', 'In Progress', 45, '2026-07-13', '2026-05-23', '[]'],
      [38, 'Drainage line layout design — Zone 8', 'Review slope specifications and finalize storm water line routes.', 'Green Valley Residency', 'anil', 'rajeev', 'Medium', 'Completed', 100, '2026-06-03', '2026-05-08', '[]'],
      [39, 'Drywall framing inspection — Level 6', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Sunrise Mall', 'anil', 'rajeev', 'Medium', 'Completed', 100, '2026-06-21', '2026-05-29', '["{\\\"username\\\": \\\"anil\\\", \\\"author\\\": \\\"Anil\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"31 may, 12:00 am\\\"}"]'],
      [40, 'Site grading and excavation — Zone 7', 'Supervise excavator operations and confirm final grade level compliance.', 'Metro Bridge Overpass', 'anil', 'rajeev', 'High', 'In Progress', 34, '2026-07-01', '2026-05-08', '["{\\\"username\\\": \\\"anil\\\", \\\"author\\\": \\\"Anil\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"15 may, 12:00 am\\\"}"]'],
      [41, 'Rebar cage inspection — Column B5', 'Verify steel reinforcement diameters and spacing for basement Column B5.', 'Sunrise Mall', 'anil', 'rajeev', 'High', 'Pending', 0, '2026-07-10', '2026-05-16', '[]'],
      [42, 'Topographical survey verification — Phase 1', 'Conduct total station survey to verify boundary limits and contours.', 'Skyline Towers', 'anil', 'rajeev', 'High', 'In Progress', 31, '2026-07-05', '2026-05-15', '["{\\\"username\\\": \\\"anil\\\", \\\"author\\\": \\\"Anil\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"22 may, 12:00 am\\\"}"]'],
      [43, 'Site grading and excavation — Zone 3', 'Supervise excavator operations and confirm final grade level compliance.', 'Sunrise Mall', 'shabnam', 'rajeev', 'High', 'Completed', 100, '2026-06-28', '2026-05-10', '[]'],
      [44, 'Safety helmet and harness inventory', 'Conduct regular inventory check of safety equipment and replace defective gear.', 'Sunrise Mall', 'shabnam', 'rajeev', 'High', 'In Progress', 47, '2026-05-31', '2026-05-08', '["{\\\"username\\\": \\\"shabnam\\\", \\\"author\\\": \\\"Shabnam\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"17 may, 12:00 am\\\"}"]'],
      [45, 'Scaffolding load compliance audit', 'Inspect base plates, ledgers, and load capacity of active scaffolding.', 'Skyline Towers', 'shabnam', 'rajeev', 'High', 'Completed', 100, '2026-07-10', '2026-05-18', '[]'],
      [46, 'Water supply header installation — Zone 2', 'Inspect PPR pipe welding and pressure testing of header lines.', 'Green Valley Residency', 'shabnam', 'rajeev', 'High', 'Pending', 0, '2026-06-27', '2026-05-23', '[]'],
      [47, 'Basement flooring screed curing — Zone 2', 'Supervise floor hardener application and spray curing of the screed slab.', 'Green Valley Residency', 'shabnam', 'rajeev', 'Low', 'Pending', 0, '2026-06-15', '2026-05-21', '[]'],
      [48, 'Drywall framing inspection — Level 9', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Metro Bridge Overpass', 'shabnam', 'rajeev', 'Medium', 'In Progress', 50, '2026-06-14', '2026-05-09', '[]'],
      [49, 'Fine sand sieve analysis', 'Perform silt content and sieve analysis tests on incoming sand stock.', 'Sunrise Mall', 'shabnam', 'rajeev', 'Medium', 'Pending', 0, '2026-06-20', '2026-05-02', '[]'],
      [50, 'Main panel board terminal check — Block 2', 'Check torque on line wire connections and ground resistance values.', 'Skyline Towers', 'shabnam', 'rajeev', 'Low', 'In Progress', 74, '2026-06-07', '2026-05-10', '["{\\\"username\\\": \\\"shabnam\\\", \\\"author\\\": \\\"Shabnam\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"12 may, 12:00 am\\\"}"]'],
      [51, 'Concrete compressive strength test — Batch 6', 'Perform slump test and prepare concrete cylinder molds for laboratory testing.', 'Metro Bridge Overpass', 'shabnam', 'rajeev', 'High', 'Completed', 100, '2026-07-21', '2026-05-28', '["{\\\"username\\\": \\\"shabnam\\\", \\\"author\\\": \\\"Shabnam\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"07 jun, 12:00 am\\\"}"]'],
      [52, 'Cement batch delivery verification', 'Inspect material quality certificates and record batch numbers on arrival.', 'Skyline Towers', 'shabnam', 'rajeev', 'Low', 'Under Review', 83, '2026-06-11', '2026-05-06', '[]'],
      [53, 'Topographical survey verification — Phase 9', 'Conduct total station survey to verify boundary limits and contours.', 'Green Valley Residency', 'varsha', 'rajeev', 'Medium', 'Under Review', 86, '2026-07-04', '2026-05-24', '[]'],
      [54, 'Cement batch delivery verification', 'Inspect material quality certificates and record batch numbers on arrival.', 'Sunrise Mall', 'varsha', 'rajeev', 'Medium', 'Pending', 0, '2026-07-03', '2026-05-04', '[]'],
      [55, 'Drywall framing inspection — Level 5', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Skyline Towers', 'varsha', 'rajeev', 'High', 'Under Review', 84, '2026-06-27', '2026-05-22', '[]'],
      [56, 'Tile layout alignment check — Lobby 9', 'Review tile spacing, layout consistency, and grout line widths.', 'Sunrise Mall', 'varsha', 'rajeev', 'Low', 'Pending', 0, '2026-05-29', '2026-05-05', '[]'],
      [57, 'Safety helmet and harness inventory', 'Conduct regular inventory check of safety equipment and replace defective gear.', 'Skyline Towers', 'varsha', 'rajeev', 'Medium', 'In Progress', 65, '2026-05-28', '2026-05-06', '[]'],
      [58, 'Site grading and excavation — Zone 6', 'Supervise excavator operations and confirm final grade level compliance.', 'Green Valley Residency', 'varsha', 'rajeev', 'Low', 'In Progress', 23, '2026-07-07', '2026-05-13', '["{\\\"username\\\": \\\"varsha\\\", \\\"author\\\": \\\"Varsha\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"20 may, 12:00 am\\\"}"]'],
      [59, 'Main panel board terminal check — Block 3', 'Check torque on line wire connections and ground resistance values.', 'Green Valley Residency', 'varsha', 'rajeev', 'High', 'In Progress', 62, '2026-06-02', '2026-05-02', '["{\\\"username\\\": \\\"varsha\\\", \\\"author\\\": \\\"Varsha\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"08 may, 12:00 am\\\"}"]'],
      [60, 'Fine sand sieve analysis', 'Perform silt content and sieve analysis tests on incoming sand stock.', 'Metro Bridge Overpass', 'varsha', 'rajeev', 'High', 'Pending', 0, '2026-06-05', '2026-05-14', '[]'],
      [61, 'Underground conduit piping — Floor 4', 'Supervise PVC conduits installation for electrical power lines.', 'Green Valley Residency', 'varsha', 'rajeev', 'Medium', 'Under Review', 89, '2026-07-01', '2026-05-28', '[]'],
      [62, 'Formwork stability check — Slab 7', 'Inspect shoring, bracing, and prop alignments before concrete pouring.', 'Metro Bridge Overpass', 'varsha', 'rajeev', 'Medium', 'Pending', 0, '2026-07-02', '2026-05-26', '[]'],
      [63, 'Exterior glass facade installation — Block 9', 'Coordinate structural glazing brackets installation and safety netting.', 'Sunrise Mall', 'gopal', 'rajeev', 'Low', 'Under Review', 80, '2026-06-10', '2026-05-05', '[]'],
      [64, 'Slab curing compound application — Level 1', 'Monitor wet burlap curing duration and apply curing compound on Slab 1.', 'Skyline Towers', 'gopal', 'rajeev', 'Low', 'Completed', 100, '2026-06-22', '2026-05-13', '["{\\\"username\\\": \\\"gopal\\\", \\\"author\\\": \\\"Gopal\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"22 may, 12:00 am\\\"}"]'],
      [65, 'Topographical survey verification — Phase 7', 'Conduct total station survey to verify boundary limits and contours.', 'Green Valley Residency', 'gopal', 'rajeev', 'Medium', 'Pending', 0, '2026-07-10', '2026-05-24', '[]'],
      [66, 'Geotechnical soil testing — Phase 9', 'Perform borehole drilling and soil compaction tests for Phase 9 foundations.', 'Green Valley Residency', 'gopal', 'rajeev', 'Medium', 'Completed', 100, '2026-06-14', '2026-05-04', '["{\\\"username\\\": \\\"gopal\\\", \\\"author\\\": \\\"Gopal\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"06 may, 12:00 am\\\"}"]'],
      [67, 'Scaffolding load compliance audit', 'Inspect base plates, ledgers, and load capacity of active scaffolding.', 'Metro Bridge Overpass', 'gopal', 'rajeev', 'Low', 'Completed', 100, '2026-06-26', '2026-05-12', '["{\\\"username\\\": \\\"gopal\\\", \\\"author\\\": \\\"Gopal\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"21 may, 12:00 am\\\"}"]'],
      [68, 'Drainage line layout design — Zone 4', 'Review slope specifications and finalize storm water line routes.', 'Sunrise Mall', 'gopal', 'rajeev', 'Low', 'Completed', 100, '2026-06-23', '2026-05-23', '["{\\\"username\\\": \\\"gopal\\\", \\\"author\\\": \\\"Gopal\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"28 may, 12:00 am\\\"}"]'],
      [69, 'Basement flooring screed curing — Zone 9', 'Supervise floor hardener application and spray curing of the screed slab.', 'Skyline Towers', 'gopal', 'rajeev', 'Medium', 'Under Review', 86, '2026-07-11', '2026-05-15', '["{\\\"username\\\": \\\"gopal\\\", \\\"author\\\": \\\"Gopal\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"21 may, 12:00 am\\\"}"]'],
      [70, 'Tile layout alignment check — Lobby 8', 'Review tile spacing, layout consistency, and grout line widths.', 'Sunrise Mall', 'gopal', 'rajeev', 'Low', 'In Progress', 75, '2026-06-16', '2026-05-17', '["{\\\"username\\\": \\\"gopal\\\", \\\"author\\\": \\\"Gopal\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"22 may, 12:00 am\\\"}"]'],
      [71, 'Water supply header installation — Zone 6', 'Inspect PPR pipe welding and pressure testing of header lines.', 'Skyline Towers', 'gopal', 'rajeev', 'High', 'Under Review', 87, '2026-06-28', '2026-05-27', '[]'],
      [72, 'Site grading and excavation — Zone 4', 'Supervise excavator operations and confirm final grade level compliance.', 'Sunrise Mall', 'gopal', 'rajeev', 'Low', 'Pending', 0, '2026-07-01', '2026-05-16', '[]'],
      [73, 'Cement batch delivery verification', 'Inspect material quality certificates and record batch numbers on arrival.', 'Sunrise Mall', 'nisha', 'rajeev', 'Medium', 'Completed', 100, '2026-06-07', '2026-05-09', '["{\\\"username\\\": \\\"nisha\\\", \\\"author\\\": \\\"Nisha\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"10 may, 12:00 am\\\"}"]'],
      [74, 'Waste debris clearance coordination', 'Organize waste collection vehicles and record disposal manifesting.', 'Sunrise Mall', 'nisha', 'rajeev', 'High', 'In Progress', 69, '2026-06-27', '2026-05-03', '[]'],
      [75, 'Topographical survey verification — Phase 8', 'Conduct total station survey to verify boundary limits and contours.', 'Green Valley Residency', 'nisha', 'rajeev', 'Medium', 'Under Review', 94, '2026-07-12', '2026-05-21', '["{\\\"username\\\": \\\"nisha\\\", \\\"author\\\": \\\"Nisha\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"30 may, 12:00 am\\\"}"]'],
      [76, 'Tile layout alignment check — Lobby 3', 'Review tile spacing, layout consistency, and grout line widths.', 'Sunrise Mall', 'nisha', 'rajeev', 'Medium', 'Under Review', 87, '2026-07-27', '2026-05-28', '[]'],
      [77, 'Water supply header installation — Zone 8', 'Inspect PPR pipe welding and pressure testing of header lines.', 'Green Valley Residency', 'nisha', 'rajeev', 'Medium', 'Completed', 100, '2026-06-11', '2026-05-04', '[]'],
      [78, 'Internal plastering audit — Floor 6', 'Inspect surface smoothness and check plaster thickness for internal walls.', 'Skyline Towers', 'nisha', 'rajeev', 'High', 'In Progress', 39, '2026-06-12', '2026-05-14', '["{\\\"username\\\": \\\"nisha\\\", \\\"author\\\": \\\"Nisha\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"16 may, 12:00 am\\\"}"]'],
      [79, 'Drywall framing inspection — Level 7', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Metro Bridge Overpass', 'nisha', 'rajeev', 'Low', 'Completed', 100, '2026-06-07', '2026-05-15', '[]'],
      [80, 'Basement flooring screed curing — Zone 1', 'Supervise floor hardener application and spray curing of the screed slab.', 'Sunrise Mall', 'nisha', 'rajeev', 'Medium', 'Pending', 0, '2026-06-21', '2026-05-13', '[]'],
      [81, 'HVAC duct installation — Level 9', 'Verify hanger spacing, sleeve layout, and damper positions for ducts.', 'Green Valley Residency', 'nisha', 'rajeev', 'Medium', 'In Progress', 44, '2026-07-05', '2026-05-15', '[]'],
      [82, 'Fine sand sieve analysis', 'Perform silt content and sieve analysis tests on incoming sand stock.', 'Green Valley Residency', 'nisha', 'rajeev', 'Medium', 'In Progress', 13, '2026-07-10', '2026-05-14', '["{\\\"username\\\": \\\"nisha\\\", \\\"author\\\": \\\"Nisha\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"15 may, 12:00 am\\\"}"]'],
      [83, 'Boundary wall layout marking — Zone 7', 'Set out coordinates and install marking stakes for the boundary wall.', 'Green Valley Residency', 'ahmad', 'rajeev', 'Medium', 'In Progress', 16, '2026-06-23', '2026-05-10', '[]'],
      [84, 'Tile layout alignment check — Lobby 6', 'Review tile spacing, layout consistency, and grout line widths.', 'Metro Bridge Overpass', 'ahmad', 'rajeev', 'Medium', 'Under Review', 93, '2026-06-04', '2026-05-10', '["{\\\"username\\\": \\\"ahmad\\\", \\\"author\\\": \\\"Ahmad\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"19 may, 12:00 am\\\"}"]'],
      [85, 'Site grading and excavation — Zone 6', 'Supervise excavator operations and confirm final grade level compliance.', 'Green Valley Residency', 'ahmad', 'rajeev', 'Low', 'Pending', 0, '2026-06-17', '2026-05-26', '[]'],
      [86, 'Waste debris clearance coordination', 'Organize waste collection vehicles and record disposal manifesting.', 'Green Valley Residency', 'ahmad', 'rajeev', 'High', 'Pending', 0, '2026-06-19', '2026-05-21', '[]'],
      [87, 'Concrete compressive strength test — Batch 3', 'Perform slump test and prepare concrete cylinder molds for laboratory testing.', 'Sunrise Mall', 'ahmad', 'rajeev', 'Low', 'Pending', 0, '2026-06-22', '2026-05-20', '[]'],
      [88, 'Safety helmet and harness inventory', 'Conduct regular inventory check of safety equipment and replace defective gear.', 'Metro Bridge Overpass', 'ahmad', 'rajeev', 'High', 'Pending', 0, '2026-06-25', '2026-05-26', '[]'],
      [89, 'Precast beam launching — Span 2', 'Supervise crane positioning and rigging for precast girder alignment.', 'Skyline Towers', 'ahmad', 'rajeev', 'Medium', 'Completed', 100, '2026-06-15', '2026-05-14', '[]'],
      [90, 'Cement batch delivery verification', 'Inspect material quality certificates and record batch numbers on arrival.', 'Skyline Towers', 'ahmad', 'rajeev', 'Low', 'Under Review', 83, '2026-07-22', '2026-05-27', '["{\\\"username\\\": \\\"ahmad\\\", \\\"author\\\": \\\"Ahmad\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"02 jun, 12:00 am\\\"}"]'],
      [91, 'HVAC duct installation — Level 7', 'Verify hanger spacing, sleeve layout, and damper positions for ducts.', 'Metro Bridge Overpass', 'ahmad', 'rajeev', 'High', 'Under Review', 80, '2026-07-14', '2026-05-29', '["{\\\"username\\\": \\\"ahmad\\\", \\\"author\\\": \\\"Ahmad\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"31 may, 12:00 am\\\"}"]'],
      [92, 'Drywall framing inspection — Level 6', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Sunrise Mall', 'ahmad', 'rajeev', 'Low', 'In Progress', 65, '2026-06-29', '2026-05-07', '["{\\\"username\\\": \\\"ahmad\\\", \\\"author\\\": \\\"Ahmad\\\", \\\"role\\\": \\\"Manager\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"12 may, 12:00 am\\\"}"]'],
      [93, 'Main panel board terminal check — Block 9', 'Check torque on line wire connections and ground resistance values.', 'Sunrise Mall', 'jitender', 'ahmad', 'Medium', 'Completed', 100, '2026-07-24', '2026-05-28', '[]'],
      [94, 'Concrete compressive strength test — Batch 2', 'Perform slump test and prepare concrete cylinder molds for laboratory testing.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'Medium', 'In Progress', 69, '2026-07-18', '2026-05-20', '["{\\\"username\\\": \\\"jitender\\\", \\\"author\\\": \\\"Jitender\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"26 may, 12:00 am\\\"}"]'],
      [95, 'Geotechnical soil testing — Phase 8', 'Perform borehole drilling and soil compaction tests for Phase 8 foundations.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'High', 'Completed', 100, '2026-06-19', '2026-05-08', '["{\\\"username\\\": \\\"jitender\\\", \\\"author\\\": \\\"Jitender\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"14 may, 12:00 am\\\"}"]'],
      [96, 'Slab curing compound application — Level 5', 'Monitor wet burlap curing duration and apply curing compound on Slab 5.', 'Skyline Towers', 'jitender', 'ahmad', 'Low', 'In Progress', 20, '2026-06-24', '2026-05-09', '["{\\\"username\\\": \\\"jitender\\\", \\\"author\\\": \\\"Jitender\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"13 may, 12:00 am\\\"}"]'],
      [97, 'Safety helmet and harness inventory', 'Conduct regular inventory check of safety equipment and replace defective gear.', 'Sunrise Mall', 'jitender', 'ahmad', 'Medium', 'Pending', 0, '2026-06-11', '2026-05-04', '[]'],
      [98, 'Basement flooring screed curing — Zone 4', 'Supervise floor hardener application and spray curing of the screed slab.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'Low', 'Under Review', 95, '2026-07-11', '2026-05-19', '[]'],
      [99, 'Scaffolding load compliance audit', 'Inspect base plates, ledgers, and load capacity of active scaffolding.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'Medium', 'Completed', 100, '2026-06-18', '2026-05-10', '[]'],
      [100, 'Topographical survey verification — Phase 4', 'Conduct total station survey to verify boundary limits and contours.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'High', 'In Progress', 34, '2026-06-27', '2026-05-08', '[]'],
      [101, 'Fire hydrant riser testing — Phase 9', 'Conduct hydrostatic pressure testing on wet riser pipe lines.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'High', 'In Progress', 47, '2026-06-21', '2026-05-09', '[]'],
      [102, 'HVAC duct installation — Level 3', 'Verify hanger spacing, sleeve layout, and damper positions for ducts.', 'Metro Bridge Overpass', 'jitender', 'ahmad', 'High', 'Pending', 0, '2026-06-26', '2026-05-19', '[]'],
      [103, 'Safety helmet and harness inventory', 'Conduct regular inventory check of safety equipment and replace defective gear.', 'Sunrise Mall', 'sahil', 'ahmad', 'High', 'Pending', 0, '2026-06-27', '2026-05-20', '[]'],
      [104, 'Underground conduit piping — Floor 8', 'Supervise PVC conduits installation for electrical power lines.', 'Sunrise Mall', 'sahil', 'ahmad', 'Medium', 'In Progress', 16, '2026-06-29', '2026-05-10', '[]'],
      [105, 'Boundary wall layout marking — Zone 7', 'Set out coordinates and install marking stakes for the boundary wall.', 'Sunrise Mall', 'sahil', 'ahmad', 'High', 'Pending', 0, '2026-06-04', '2026-05-06', '[]'],
      [106, 'Fire hydrant riser testing — Phase 5', 'Conduct hydrostatic pressure testing on wet riser pipe lines.', 'Skyline Towers', 'sahil', 'ahmad', 'High', 'Pending', 0, '2026-07-04', '2026-05-19', '[]'],
      [107, 'Main panel board terminal check — Block 4', 'Check torque on line wire connections and ground resistance values.', 'Sunrise Mall', 'sahil', 'ahmad', 'Medium', 'Completed', 100, '2026-07-07', '2026-05-11', '["{\\\"username\\\": \\\"sahil\\\", \\\"author\\\": \\\"Sahil\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"16 may, 12:00 am\\\"}"]'],
      [108, 'Site grading and excavation — Zone 2', 'Supervise excavator operations and confirm final grade level compliance.', 'Green Valley Residency', 'sahil', 'ahmad', 'Low', 'In Progress', 43, '2026-06-17', '2026-05-23', '[]'],
      [109, 'Rebar cage inspection — Column B9', 'Verify steel reinforcement diameters and spacing for basement Column B9.', 'Skyline Towers', 'sahil', 'ahmad', 'High', 'Pending', 0, '2026-07-02', '2026-05-15', '[]'],
      [110, 'Precast beam launching — Span 1', 'Supervise crane positioning and rigging for precast girder alignment.', 'Green Valley Residency', 'sahil', 'ahmad', 'Medium', 'Under Review', 94, '2026-06-07', '2026-05-04', '["{\\\"username\\\": \\\"sahil\\\", \\\"author\\\": \\\"Sahil\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"14 may, 12:00 am\\\"}"]'],
      [111, 'Fine sand sieve analysis', 'Perform silt content and sieve analysis tests on incoming sand stock.', 'Sunrise Mall', 'sahil', 'ahmad', 'High', 'In Progress', 29, '2026-06-08', '2026-05-10', '[]'],
      [112, 'Scaffolding load compliance audit', 'Inspect base plates, ledgers, and load capacity of active scaffolding.', 'Sunrise Mall', 'sahil', 'ahmad', 'High', 'Completed', 100, '2026-07-02', '2026-05-24', '["{\\\"username\\\": \\\"sahil\\\", \\\"author\\\": \\\"Sahil\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Completed deliverables uploaded.\\\", \\\"timestamp\\\": \\\"29 may, 12:00 am\\\"}"]'],
      [113, 'Water supply header installation — Zone 9', 'Inspect PPR pipe welding and pressure testing of header lines.', 'Sunrise Mall', 'anus', 'ahmad', 'Medium', 'Pending', 0, '2026-06-12', '2026-05-21', '[]'],
      [114, 'Drywall framing inspection — Level 6', 'Verify stud spacing and door frame reinforcement in partition walls.', 'Metro Bridge Overpass', 'anus', 'ahmad', 'High', 'Pending', 0, '2026-07-04', '2026-05-09', '[]'],
      [115, 'Fire hydrant riser testing — Phase 1', 'Conduct hydrostatic pressure testing on wet riser pipe lines.', 'Metro Bridge Overpass', 'anus', 'ahmad', 'Low', 'Pending', 0, '2026-06-26', '2026-05-26', '[]'],
      [116, 'Underground conduit piping — Floor 9', 'Supervise PVC conduits installation for electrical power lines.', 'Sunrise Mall', 'anus', 'ahmad', 'Medium', 'In Progress', 65, '2026-07-12', '2026-05-22', '["{\\\"username\\\": \\\"anus\\\", \\\"author\\\": \\\"Anus\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"30 may, 12:00 am\\\"}"]'],
      [117, 'Exterior glass facade installation — Block 7', 'Coordinate structural glazing brackets installation and safety netting.', 'Metro Bridge Overpass', 'anus', 'ahmad', 'Medium', 'Pending', 0, '2026-06-28', '2026-05-29', '[]'],
      [118, 'Internal plastering audit — Floor 7', 'Inspect surface smoothness and check plaster thickness for internal walls.', 'Sunrise Mall', 'anus', 'ahmad', 'Medium', 'Completed', 100, '2026-07-22', '2026-05-28', '[]'],
      [119, 'Boundary wall layout marking — Zone 6', 'Set out coordinates and install marking stakes for the boundary wall.', 'Metro Bridge Overpass', 'anus', 'ahmad', 'Medium', 'Pending', 0, '2026-07-10', '2026-05-26', '[]'],
      [120, 'Geotechnical soil testing — Phase 9', 'Perform borehole drilling and soil compaction tests for Phase 9 foundations.', 'Sunrise Mall', 'anus', 'ahmad', 'Medium', 'Pending', 0, '2026-06-30', '2026-05-08', '[]'],
      [121, 'Main panel board terminal check — Block 8', 'Check torque on line wire connections and ground resistance values.', 'Sunrise Mall', 'anus', 'ahmad', 'High', 'In Progress', 44, '2026-06-16', '2026-05-19', '["{\\\"username\\\": \\\"anus\\\", \\\"author\\\": \\\"Anus\\\", \\\"role\\\": \\\"User\\\", \\\"text\\\": \\\"Task initiated and progressing smoothly.\\\", \\\"timestamp\\\": \\\"27 may, 12:00 am\\\"}"]'],
      [122, 'Safety helmet and harness inventory', 'Conduct regular inventory check of safety equipment and replace defective gear.', 'Skyline Towers', 'anus', 'ahmad', 'High', 'In Progress', 30, '2026-07-05', '2026-05-11', '[]']
];
    // Map of seeded usernames to emails to dynamic inject email column values and avoid data shifts
    var userEmailMap = {};
    users.forEach(function(u) {
      userEmailMap[u[0]] = u[9] || ''; // Username is index 0, Email is index 9
    });

    tasks.forEach(function(row) {
      var assignee = row[4];
      var email = userEmailMap[assignee] || '';
      // Insert email at index 5 to keep the task rows aligned with the database schema
      var newRow = row.slice(0, 5).concat([email]).concat(row.slice(5));
      taskSheet.appendRow(newRow);
    });
  }

  // --- Ensure Attachments sheet ---
  var attachmentsSheet = getOrCreateSheet(ss, 'Attachments', [
    'AttachId', 'TaskId', 'FileName', 'FileType', 'FileUrl', 'FileId', 'ThumbnailUrl', 'UploadedBy', 'UploadedAt'
  ]);

  // --- Ensure Logs sheet ---
  var logSheet = getOrCreateSheet(ss, 'Logs', [
    'Timestamp', 'TaskId', 'FileName', 'FileUrl', 'UploadedBy'
  ]);

  // --- Ensure Sessions and Settings sheets ---
  var sessionSheet = getOrCreateSheet(ss, 'Sessions', ['Timestamp', 'Username', 'Action', 'Details']);
  var settingsSheet = getOrCreateSheet(ss, 'Settings', ['Setting', 'Value', 'Description', 'LastModified']);

  // Seed default settings if empty
  if (settingsSheet.getLastRow() < 2) {
    var defaultSettings = [
      ['showLogsToAdmins', 'true', 'Allow Admin users to access the Activity Logs page.'],
      ['showAttachmentsToAdmins', 'true', 'Allow Admin users to access the Snapshots page.'],
      ['showTeamToAdmins', 'true', 'Allow Admin users to access the Team Directory page.'],
      ['showProjectsToAdmins', 'true', 'Allow Admin users to access the Projects page.'],
      ['showSettingsToAdmins', 'true', 'Allow Admin users to open the Settings page.'],
      ['showLogsToManagers', 'false', 'Let Project Managers check system activity and logs.'],
      ['showAttachmentsToManagers', 'false', 'Let Project Managers browse the snapshots gallery.'],
      ['showTeamToManagers', 'false', 'Let Project Managers view the Team Directory page.'],
      ['showProjectsToManagers', 'false', 'Let Project Managers view the Projects overview page.'],
      ['showSettingsToManagers', 'false', 'Allow Project Managers to open the Settings page.'],
      ['allowNewTaskToManagers', 'true', 'Let Project Managers create new tasks using the +New Task button.'],
      ['showLogsToUsers', 'false', 'Let team members check system upload activities.'],
      ['showAttachmentsToUsers', 'false', 'Let team members browse the snapshots gallery.'],
      ['showTeamToUsers', 'false', 'Let team members view the Team Directory page.'],
      ['showProjectsToUsers', 'false', 'Let team members view the Projects overview page.'],
      ['showSettingsToUsers', 'false', 'Let team members open the Settings page if enabled.'],
      ['allowNewTaskToUsers', 'true', 'Let team members create new tasks using the +New Task button.'],
      ['disableDemoSeed', 'false', 'Prevent demo data from being re-seeded.']
    ];
    defaultSettings.forEach(function(row) { settingsSheet.appendRow(row); });
  }

  // Format all sheets consistently
  [userSheet, projectSheet, taskSheet, attachmentsSheet, logSheet, sessionSheet, settingsSheet].forEach(formatSheet);

  return 'Default data seeded successfully!';
}

function formatSheet(sheet) {
  if (!sheet) return;
  sheet.setFrozenRows(1);
  var maxCols = sheet.getMaxColumns();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var numRows = values.length;
  var numCols = numRows > 0 ? values[0].length : 0;
  if (numRows === 0 || numCols === 0) return;

  // Remove all extra columns after the last defined header in row 1
  var headers = values[0];
  var lastHeaderCol = 0;
  for (var c = 0; c < maxCols; c++) {
    var headerValue = c < headers.length ? headers[c] : '';
    if (String(headerValue || '').trim() !== '') {
      lastHeaderCol = c + 1;
    }
  }
  if (lastHeaderCol > 0 && lastHeaderCol < maxCols) {
    sheet.deleteColumns(lastHeaderCol + 1, maxCols - lastHeaderCol);
    numCols = lastHeaderCol;
    range = sheet.getRange(1, 1, numRows, numCols);
  }

  // Keep header row centered, data rows aligned by type
  sheet.getRange(1, 1, 1, numCols).setHorizontalAlignment('center');

  var alignments = [];
  for (var r = 1; r < numRows; r++) {
    var rowAlign = [];
    for (var c = 0; c < numCols; c++) {
      var cellValue = values[r][c];
      var alignment = 'left';
      if (cellValue instanceof Date || typeof cellValue === 'number') {
        alignment = 'center';
      } else if (typeof cellValue === 'string') {
        var trimmed = cellValue.trim();
        if (/^\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2}:\d{2})?$/.test(trimmed) || /^\d+(\.\d+)?$/.test(trimmed)) {
          alignment = 'center';
        }
      }
      rowAlign.push(alignment);
    }
    alignments.push(rowAlign);
  }

  if (alignments.length > 0) {
    sheet.getRange(2, 1, alignments.length, numCols).setHorizontalAlignments(alignments);
  }

  range.setVerticalAlignment('middle')
       .setFontFamily('Calibri')
       .setFontSize(10);

  // Auto-fit all active columns to their content
  for (var col = 1; col <= numCols; col++) {
    sheet.autoResizeColumn(col);
  }
}

function formatAllSheets() {
  const ss = getSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(formatSheet);
  return 'All sheets formatted successfully!';
}

function sendTestMail() {
  try {
    var email = 'automationbyanil@gmail.com';
    var subject = 'Aarcon — Test Email';
    var textBody = 'This is a test email from Aarcon Task Manager.\n\nIf you received this, email notifications are working correctly.\n\n— Aarcon System';
    var htmlBody = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Arial,Helvetica,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8;padding:30px 0;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">
      <tr>
        <td style="background:linear-gradient(135deg,#1a2540 0%,#0f1829 100%);padding:32px 36px 28px;text-align:center;">
          <div style="display:inline-block;background:#f59e0b;color:#1a1a1a;font-weight:700;font-size:13px;letter-spacing:1px;padding:5px 12px;border-radius:20px;text-transform:uppercase;">⚙️ AARCON</div>
          <h1 style="color:#ffffff;margin:14px 0 4px;font-size:22px;font-weight:700;">✅ Test Email</h1>
          <p style="color:#8fa3c0;margin:0;font-size:13px;">This is a test email from Aarcon Task Manager.</p>
        </td>
      </tr>
      <tr>
        <td style="background:#ffffff;padding:32px 36px;">
          <div style="background:#e8f5e9;border:1px solid #c8e6c9;border-radius:10px;padding:20px 22px;text-align:center;">
            <p style="margin:0 0 8px;font-size:16px;font-weight:700;color:#2e7d32;">✓ Email Notifications Working</p>
            <p style="margin:0;font-size:13px;color:#555;">If you received this email, Aarcon's email system is configured correctly and ready to send task notifications.</p>
          </div>
        </td>
      </tr>
      <tr>
        <td style="background:#f7f9fc;border-top:1px solid #e4e9f0;padding:18px 36px;text-align:center;">
          <p style="margin:0;font-size:12px;color:#8fa3c0;">© 2026 <strong>Aarcon</strong> — Construction Task Manager</p>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</body>
</html>`;

    MailApp.sendEmail(email, subject, textBody, {
      htmlBody: htmlBody,
      name: 'Aarcon Task Manager',
      replyTo: 'automationbyanil@gmail.com'
    });

    return 'Test email sent successfully to ' + email + '!';
  } catch(e) {
    return 'Failed to send test email: ' + e.message;
  }
}
