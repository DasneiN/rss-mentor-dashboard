const fs = require('fs');
const XLSX = require('xlsx');
const _ = require('lodash');

function saveFile (fileName, fileContext) {
  fs.writeFile(fileName, fileContext, (err) => {
    if (err) {
      console.log(`Error: ${err}`);
    } else {
      console.log(`File ${fileName} was saved!`);
    }
  });
}

function updateTasks() {
  const tasks = [];
  
  const workbook = XLSX.readFile('./xlsx/Tasks.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref']);

  for (let rowNum = 1; rowNum <= range.e.r; rowNum++) {
    const taskName = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 0 })];
    const taskLink = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })] || {v: ''};
    const taskStatus = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 2 })];

    tasks.push({
      'name': taskName.v,
      'link': taskLink.v,
      'status': taskStatus.v,
    });
  }
  
  return tasks;
}

function updateMentorsStudentsPairs() {
  const workbook = XLSX.readFile('./xlsx/Mentor-students pairs.xlsx');

  const sheet_1 = workbook.Sheets[workbook.SheetNames[0]];
  const range_1 = XLSX.utils.decode_range(sheet_1['!ref']);

  const sheet_2 = workbook.Sheets[workbook.SheetNames[1]];
  const range_2 = XLSX.utils.decode_range(sheet_2['!ref']);
  
  const pairs = [];

  for (let rowNum = 1; rowNum <= range_1.e.r; rowNum++) {
    if (sheet_1[XLSX.utils.encode_cell({ r: rowNum, c: 0 })]) {
      const mentorName = sheet_1[XLSX.utils.encode_cell({ r: rowNum, c: 0 })];
      const studentGit = clearGitLink(sheet_1[XLSX.utils.encode_cell({ r: rowNum, c: 1 })].v);
      
      pairs.push({
        mentor: mentorName.v,
        student: studentGit,
      });
    }
  }
  
  const groupedPairs = _.groupBy(pairs, 'mentor');
  
  const mentors = [];
  
  for (let rowNum = 1; rowNum <= range_2.e.r; rowNum++) {
    if (sheet_2[XLSX.utils.encode_cell({ r: rowNum, c: 0 })]) {
      const mentorName = sheet_2[XLSX.utils.encode_cell({ r: rowNum, c: 0 })].v;
      const mentorSurname = sheet_2[XLSX.utils.encode_cell({ r: rowNum, c: 1 })].v;
      const mentorFullname = `${mentorName} ${mentorSurname}`;
      
      const mentorGitLink = sheet_2[XLSX.utils.encode_cell({ r: rowNum, c: 4 })];
      const mentorGit = clearGitLink(mentorGitLink.v);
      
      mentors.push({
        name: mentorName,
        surname: mentorSurname,
        git: mentorGit,
        students: groupedPairs[mentorFullname].map(v => {
          return {
            student: v.student,
            tasks: [],
          }
        }),
      });
    }
  }

  return updateScore(mentors);
}

function updateScore(mentors) {
  const workbook = XLSX.readFile('./xlsx/Mentor score.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref']);

  let errors = [];
  
  for (let rowNum = 1; rowNum <= range.e.r; rowNum++) {
    if (sheet[XLSX.utils.encode_cell({ r: rowNum, c: 0 })]) {
      const mentorGit = clearGitLink(sheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })].v);
      const studentGit = clearGitLink(sheet[XLSX.utils.encode_cell({ r: rowNum, c: 2 })].v);
      const taskName = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 3 })].v;

      
      try {
        let mentorObj = _.find(mentors, (v) => _.find(v.students, { student: studentGit }));

        if (!mentorObj) {
          mentorObj = _.find(mentors, { git: mentorGit });
          
          mentorObj.students.push({
            student: studentGit,
            tasks: [],
          });
        }
        
        const studentObj = _.find(mentorObj.students, { student: studentGit });
        studentObj.tasks.push(taskName);
      } catch (error) {
        console.log('===============================================');
        console.log(`Mentor Git: ${mentorGit}`);
        console.log(`Student Git: ${studentGit}`);
        console.log(`Task: ${taskName}`);
        console.log(`Error: ${error}`);
        console.log('===============================================');

        errors.push(error);
      }
    }
  }
  
  console.log('Errors: ', _.size(errors));
  return mentors;
}

function clearGitLink(link) {
  return link
    .toString()
    .split('github.com/')
    .slice(-1)
    .join('')
    .toLowerCase()
    .replace('-2018q3', '')
    .replace('rolling-scopes-school/', '')
    .replace('/', '');
}

const newData = {
  tasks: updateTasks(),
  mentors: updateMentorsStudentsPairs(),
}

saveFile('json/data.json', JSON.stringify(newData));
