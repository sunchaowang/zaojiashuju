import xlsx from 'xlsx/dist/xlsx.core.min';
import moment from "moment";

function sheet2blob(sheet, sheetName) {
  sheetName = sheetName || 'sheet1';
  var workbook = {
    SheetNames: [sheetName],
    Sheets: {}
  };
  workbook.Sheets[sheetName] = sheet; // 生成excel的配置项

  const wopts = {
    bookType: 'xlsx', // 要生成的文件类型
    bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
    type: 'binary'
  };
  const wbout = xlsx.write(workbook, wopts);
  const blob = new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }); // 字符串转ArrayBuffer
  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  return blob;
}

function openDownloadDialog(url, saveName) {
  if (typeof url == 'object' && url instanceof Blob) {
    url = URL.createObjectURL(url); // 创建blob地址
  }
  const aLink = document.createElement('a');
  aLink.href = url;
  aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
  let event;
  if (window.MouseEvent) event = new MouseEvent('click');
  else {
    event = document.createEvent('MouseEvents');
    event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
  }
  aLink.dispatchEvent(event);
}
Date.prototype.Format = function (fmt) { //author: meizz
  const o = {
    "M+": this.getMonth() + 1, //月份
    "d+": this.getDate(), //日
    "h+": this.getHours(), //小时
    "m+": this.getMinutes(), //分
    "s+": this.getSeconds(), //秒
    "q+": Math.floor((this.getMonth() + 3) / 3), //季度
    "S": this.getMilliseconds() //毫秒
  };
  if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
  for (let k in o)
    if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
  return fmt;
}

const header = ["id", "serial", "ip", "terminal", "device"];
const headerDisplay = {
  id: "序号",
  serial: "流水号",
  ip: "IP",
  terminal: "终端",
  device: "设备",
  createTime: "访问时间"
}


function downloadExcel(cellData, fileName) {
  openDownloadDialog(sheet2blob(xlsx.utils.json_to_sheet([
    headerDisplay,
    ...cellData
  ], {
    header,
    skipHeader: true
  })), fileName);
}

const terminals = [
  "PC端", "移动端"
]
const devices = [
  "苹果", "安卓", "pad"
]

// window.document.querySelector(".downloadExcel").addEventListener("click", function () {
//
// })

const App = {
  data() {
    return {
      form: {
        time: [],
        dateStart: "",
        dateEnd: ""
      },
    };
  },
  methods: {
    downloadClick() {
      const cellData = []
      if (this.form.time.length < 2) {
        return this.$message.warning("选择时间");
      }
      const timeStart = this.form.time[0];
      const timeEnd = this.form.time[1];

      if (moment(timeEnd).month() !== moment(timeStart).month() ) {
        return this.$message.warning("不能跨月");
      }

      const month = moment(timeEnd).month() + 1;
      const total = this.form.total;
      const year = moment(timeEnd).year();
      const dateStart = moment(timeStart).date();
      const dateEnd = moment(timeEnd).date();

      let date = moment().year(year).month(Number(month) - 1)
      const daysInMonth = date.daysInMonth()
      let days = [];
      // if (daysInMonth < Number(dateEnd)) {
      //
      // }else {
      //   days = [...new Array(daysInMonth)].map((row, index) => {
      //     return index + 1;
      //   })
      // }
      for (let i = Number(dateStart); i <= Number(dateEnd); i++) {
        days.push(Number(i));
      }
      const filename = '点击数量汇总_' + '_' + year + '_' + month+ '_' + total + '_' + dateStart+ '_' + dateEnd +'.xlsx'
      // return;
      let index = 0;
      do{
        date = moment().year(year).month(Number(month) - 1)
        const terminal = terminals[Math.floor(Math.random() * terminals.length)];
        const randomIp = () => Array(4).fill(0).map((_, i) => Math.floor(Math.random() * 255) + (i === 0 ? 1 : 0)).join('.');

        let device = "";
        if (terminal == "移动端") {
          device = devices[Math.floor(Math.random() * devices.length)];
        }
        date.date(days[Math.floor(Math.random() * days.length)]);
        date.hour( Math.random() * 23);
        date.minute(Math.random() * 59);
        date.second( Math.random() * 59);
        cellData.push({
          id: index +1,
          serial: date.format("yyyyMMDDHHmmss" + Math.floor((Math.random() * 9 + 1) * 100000)),
          ip: randomIp(),
          terminal,
          device,
          createTime: date.format("yyyy.MM.DD HH:mm:ss"),
        } )
        index++;
      }
      while(index<total)

      setTimeout(function () {
        downloadExcel(cellData, filename)
      },2)
    }
  }
};
const app = Vue.createApp(App);
app.use(ElementPlus);
app.mount("#app");
