const express = require("express");
const bodyParser = require("body-parser");
const dayjs = require("dayjs");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const axios = require("axios");
const path = require("path");
const app = express();
const port = 3000;

app.set("view engine", "ejs");
app.use(bodyParser.urlencoded({ extended: false }));
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + path.extname(file.originalname)); // đặt tên file
  },
});
const upload = multer({ storage: storage });

app.get("/", (req, res) => {
  res.render("index");
});

app.get("/success", (req, res) => {
  res.render("success");
});

app.get("/fail", (req, res) => {
  res.render("fail");
});

app.post("/data", upload.single("myfile"), async (req, res) => {
  try {
    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets["PL2_Theo_Doi_SL_Ngay_2025"];
    const data = [];
    const range = xlsx.utils.decode_range(worksheet["!ref"]);
    const ctt = getValueService("Cổng thanh toán", 12, worksheet, range);
    const thd = getValueService("Thu hộ điện", 26, worksheet, range);
    const thn = getValueService("Thu hộ nước", 27, worksheet, range);
    const thtcbh = getValueService(
      "Thu hộ tài chính bảo hiểm",
      29,
      worksheet,
      range
    );
    data.push(...ctt, ...thd, ...thn, ...thtcbh);
    fs.unlinkSync(filePath);
    for (const item of data) {
      await callApiSheet(item);
    }
    res.redirect("/success");
  } catch (error) {
    console.log(error);
    res.redirect("/fail");
  }
});

app.post("/service", async (req, res) => {
  try {
    const startDate = dayjs(req.body.startDate);
    const endDate = dayjs(req.body.endDate);
    const diffDays = endDate.diff(startDate, "day") + 1;
    const payLoad = [];
    for (let i = 0; i < diffDays; i++) {
      const currentDay = startDate.add(i, "day");
      payLoad.push({
        dichVu: req.body.dichVu,
        thucHien: (parseInt(req.body.thucHien) / diffDays).toFixed(2),
        ngay: currentDay.format("MM/DD/YYYY"),
        thangNam: currentDay.format("MM/YYYY"),
      });
    }
    for (const item of payLoad) {
      await callApiSheet(item);
    }
    res.redirect("/success");
  } catch (error) {
    console.log(error);
    res.redirect("/fail");
  }
});

app.post("/subscription", async (req, res) => {
  try {
    const date = dayjs(req.body.date);
    await callApiSheet2({
      dichVu: req.body.dichVu,
      thucHien: parseInt(req.body.thucHien),
      ngay: date.format("MM/DD/YYYY"),
      thangNam: date.format("MM/YYYY"),
    });
    res.redirect("/success");
  } catch (error) {
    console.log(error);
    res.redirect("/fail");
  }
});

const formatDate = (value, type) => {
  const str = value.toString();
  const year = str.substring(0, 4);
  const month = str.substring(4, 6);
  const day = str.substring(6, 8);
  if (type == "date") {
    return `${month}/${day}/${year}`;
  } else {
    return `${month}/${year}`;
  }
};

const callApiSheet = async (data) => {
  let config = {
    method: "post",
    maxBodyLength: Infinity,
    url: "https://script.google.com/macros/s/AKfycbz9bFFpFDVxPkREZMrLLfueEr7kFVpetws8tAb71HzRzkGCwV4E3FhFp6zq2Ql4dhA/exec",
    headers: {
      "Content-Type": "application/json",
    },
    data: JSON.stringify(data),
  };
  await axios
    .request(config)
    .then((response) => {
      console.log(JSON.stringify(response.data));
    })
    .catch((error) => {
      console.log(error);
    });
};

const callApiSheet2 = async (data) => {
  let config = {
    method: "post",
    maxBodyLength: Infinity,
    url: "https://script.google.com/macros/s/AKfycbwQamL_VpiX39FDnkuKvsQ4-cMR3w4hgxiW2iMmxVCBv6lVR6NH2pvTcet9veJ0U754/exec",
    headers: {
      "Content-Type": "application/json",
    },
    data: JSON.stringify(data),
  };
  await axios
    .request(config)
    .then((response) => {
      console.log(JSON.stringify(response.data));
    })
    .catch((error) => {
      console.log(error);
    });
};

const getValueService = (name, col, worksheet, range) => {
  const data = [];
  for (let row = range.s.r; row <= range.e.r; row++) {
    const cellNgay = worksheet[xlsx.utils.encode_cell({ r: row, c: 0 })];
    const cellThuhien = worksheet[xlsx.utils.encode_cell({ r: row, c: col })];
    data.push({
      dichVu: name,
      thucHien: cellThuhien
        ? parseFloat(cellThuhien.v / 1000000).toFixed(2)
        : null,
      ngay: cellNgay ? formatDate(cellNgay.v, "date") : null,
      thangNam: cellNgay ? formatDate(cellNgay.v) : null,
    });
  }
  const filtered = data.filter(
    (item) => item.thucHien !== null && item.ngay !== null
  );
  return filtered.slice(1);
};

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
