function parseGmailBodyText(text) {
    const match = text.match(/\[コース名\] : (.*)\n\[課題名\] : (.*)\n(?:\[受付開始日時\] : (.*)\n)?(?:\[受付終了日時\] : (.*)\n)?PC : (https?:\/\/[\w/:%#\$&\?\(\)~\.=\+\-]+)/);
    if (match) {
      return {
        courseName: match[1]?.trim() || null,
        taskName: match[2]?.trim() || null,
        taskStartDateTime: match[3]?.trim() || null,
        taskEndDateTime: match[4]?.trim() || null,
        url: match[5]?.trim() || null,
      };
    }
    return {
      courseName: text.match(/\[コース名\] : (.*)/)?.[1]?.trim() || null,
      taskName: text.match(/\[課題名\] : (.*)/)?.[1]?.trim() || null,
      taskStartDateTime: text.match(/\[受付開始日時\] : (.*)/)?.[1]?.trim() || null,
      taskEndDateTime: text.match(/\[受付終了日時\] : (.*)/)?.[1]?.trim() || null,
      url: text.match(/PC : (https?:\/\/[\w/:%#\$&\?\(\)~\.=\+\-]+)/)?.[1]?.trim() || null,
    };
  }
  
  function addCalendar(courseName, taskName, startDateTime, endDateTime, url) {
    if (!courseName || !taskName || !startDateTime || !endDateTime) {
      Logger.log('カレンダーへの追加に必要な情報が不足しています。');
      return;
    }
  
    const calendar = CalendarApp.getDefaultCalendar();
    const title = `【課題】${courseName} - ${taskName}`;
    const description = url ? `課題URL: ${url}` : '';
  
    try {
      calendar.createEvent(title, startDateTime, endDateTime, { description: description });
      // Logger.log(`"${title}" をカレンダーに追加しました。開始: ${startDateTime}, 終了: ${endDateTime}`);
    } catch (e) {
      Logger.log(`カレンダーへの追加に失敗しました: ${e}`);
    }
  }
  
  function main() {
    const subjectQuery = "レポート公開のお知らせ";
    const senderQuery = "from:do-not-reply@manaba.jp";
    const labelName = "新着課題";
    const halfYearAgo = new Date();
    halfYearAgo.setMonth(halfYearAgo.getMonth() - 6);
    halfYearAgo.setHours(0, 0, 0, 0); // 半年前の午前0時
  
    const searchQuery = `${subjectQuery} ${senderQuery} -label:${labelName} after:${halfYearAgo.toISOString().slice(0, 10)}`;
    const threads = GmailApp.search(searchQuery);
    const newTasksLabel = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
  
    let processedCount = 0;
  
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        const receivedDate = message.getDate();
        // 受信日が半年前より新しいかチェック
        if (receivedDate >= halfYearAgo) {
          const body = message.getPlainBody();
          const taskInfo = parseGmailBodyText(body);
  
          const getStartDateTime = () => {
            if (taskInfo.taskStartDateTime) {
              return new Date(taskInfo.taskStartDateTime);
            }
            const date = new Date(receivedDate);
            date.setHours(0, 0, 0, 0);
            return date;
          };
  
          const getEndDateTime = (start) => {
            if (taskInfo.taskEndDateTime) {
              return new Date(taskInfo.taskEndDateTime);
            }
            if (start) {
              const date = new Date(start);
              date.setDate(date.getDate() + 1);
              date.setHours(0, 0, 0, 0);
              return date;
            }
            const date = new Date(receivedDate);
            date.setDate(date.getDate() + 1);
            date.setHours(0, 0, 0, 0);
            return date;
          };
  
          const startDateTime = getStartDateTime();
          const endDateTime = getEndDateTime(startDateTime);
  
          if (taskInfo.courseName && taskInfo.taskName && startDateTime && endDateTime) {
            addCalendar(
              taskInfo.courseName,
              taskInfo.taskName,
              startDateTime,
              endDateTime,
              taskInfo.url
            );
            thread.addLabel(newTasksLabel);
            processedCount++;
          } else {
            Logger.log('カレンダーに追加するための必須情報が不足しています。');
            Logger.log(taskInfo);
            Logger.log('受信日:', receivedDate);
          }
        }
      }
    }
    Logger.log(`${processedCount} 件の課題をカレンダーに追加しました。`);
  }