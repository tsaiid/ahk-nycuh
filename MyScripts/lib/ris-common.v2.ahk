; RIS specific functions

SleepThenTab(sleepTime := 400, shiftTab := false)
{
  Sleep sleepTime
  if (shiftTab) {
    Send "+{Tab}"
  } else {
    Send "{Tab}"
  }
}
