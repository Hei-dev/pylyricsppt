# PyLyricsPPT
 A python script to create lyrics PowerPoint automatically

## Software used
 **Install these software beforehand**
 Python 3
 Python pptx (pip install python-pptx)

## Procedures
Step 1. After downloading the program/script, create a folder in the directory of the .py file.\
Step 2. Create a file named "lyrics.txt" in that folder.\
Step 3. Add the lyrics to "lyrics.txt" with the following format:
 - Indicate the start of the song with "T:"
`T:Title`
 - Indicate the lyrics (Verse) with "V:", followed by the verse lyrics (New line=New slide)
```
V:
Verse Verse Verse
Lyrics Lyrics Lyrics
```
 - Note: For multiple verses, indicate them by adding a number after "V"
```
V1:
Lyrics Lyrics
V2:
Verse Verse
```
 - Indicate the lyrics (Chorus) with "C:", followed by the Chorus lyrics (New line=New slide)
```
C:
Chorus Chorus Chorus
Lyrics Lyrics Lyrics
```
 - Note: For multiple choruses, indicate them by adding a number after "C"
 - Indicate the lyrics (Bridge) with "B:", followed by the bridge lyrics (New line=New slide)
B:
Bridge Bridge Bridge
Lyrics Lyrics Lyrics
 - Note: For multiple bridges, indicate them by adding a number after "B"
 - To indicate the start of another song, add another "T:" followed by the song title.
`T:Title2`
### Sample lyrics.txt
```
T:蒙著嘴說愛你
V1:
可否一覺甦醒 推開窗看看風景
混亂日子不見蹤影
鳥語花香呼應 阻止黑暗得逞

始終相信惡夢會過
健康與和平 不必依靠偶然性
愛接愛 人連人 可以一拼
活化 我心境

V2:
多想趕快歸隊 必須將鬥志高舉
但願上天給我應許
眷顧身邊的愛侶 巴不得吻這張嘴

即將一切會成過去
幸福悄悄伴隨 此刻身處過程裡
嚇怕了 難捱時 不撤不退
集氣 再爭取

C1:
So I say I love you 只有愛恆久不枯
生活在劫難裡 心靈從未給沾污
天給我貧病困苦 笑著去吃苦
沒怨恨誰人可惡
So I say I love you 花半秒唇齒功夫
使淡靜歲月變豐富
即使要蒙著我嘴 我亦可高呼
全憑愛令我堅持 還有你悉心照顧

B:
不在乎 多在乎
生於這世界 還有一些牽掛令我在乎
等天際亮了 不區分喜惡
除下這隔膜打個招呼

C2:
So I say I love you 只有愛恆久不枯
生活在劫難裡 希望從未給沾污
天給我磨練也好 我未敢辜負
誰要被懷疑低估

Here We say I love you 都變了甜品師傅
巧妙地化掉這點苦
即使要蒙著我嘴 更大聲歡呼
全憑愛令人堅持 還有各位的照顧

T:世上只有
V1:
望著你講 也許更易
濃於水的三個字
從我降世 一開始 到永遠 不休止
你亦是我支柱 動力和意義

V2:
但是我知 你都有夢
仍將一生給我用
全個世界 幾多種 愛與愛 在互動
也未及這種愛 能完全獻奉

C:
You make me cry make me smile
Make me feel that love is true
謝謝你的關顧 與及無償的愛護
年月漫漫 多艱苦
你也永遠優先擔心我喜惡
唯恐我並未得到 最貼身保護
Oh I love you

B:
Yes I do
I always do
```
Step 4. Create a file named "font.json". Paste the following content into the file:
```
{
  "typeface":"FONT"
  "size":{
    "Title":145,
    "Lyrics":80
  }
  "max_line_length":14
}
```
 - Change `FONT` with the font name intended to use.
 - Replace any other integer value with user's preference

Step 5(OPTIONAL). Insert background images:\
 5a. Create, for each of the songs, a folder named as the song title.\
 5b. Place images into the respective folder.\
 5c. Rename the images to codes that represent the "places" that it is intended to use\
     e.g. V1.jpg refer to use V1.jpg for the slide background of the Verse 1 lyrics\
          For codes, refer to lyrics.txt (e.g. V1,B,C,C1,etc.)\
 5d. Resize/Crop the images to the 16:9 ratio\

Step 6: Run the command: `python3 pyppt.py`\
File will be output as "test.pptx"
