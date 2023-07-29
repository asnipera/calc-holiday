<script setup lang="ts">
import { ref } from 'vue'
import type { UploadChangeParam } from 'ant-design-vue'
import { read, utils, writeFile } from 'xlsx'
import dayjs from 'dayjs'
import dayOfYear from 'dayjs/plugin/dayOfYear'
import { UploadOutlined } from '@ant-design/icons-vue'
dayjs.extend(dayOfYear)

// 四舍五入，保留一位小数
function round(num: number) {
  return Math.round(num * 10) / 10
}

let tableData: any[]
const loading = ref(false)
function download() {
  if (!tableData?.length) return
  loading.value = true
  const wb = utils.book_new()
  const ws = utils.json_to_sheet(tableData, { skipHeader: true })
  const wchs = {
    '!cols': [
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 21 }, // 第一列
      { wch: 12 }, // 第一列
      { wch: 21 }, // 第一列
      { wch: 15 }, // 第一列
      { wch: 18 }, // 第一列
      { wch: 15 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 10 }, // 第一列
      { wch: 18 }, // 第一列
      { wch: 18 } // 第一列
    ],
    '!rows': [{ hpx: 16 }]
  }
  const merges = {
    '!merges': [
      {
        s: {
          // s开始
          c: 0, // 开始列
          r: 0 // 开始取值范围
        },
        e: {
          // e结束
          c: 24, // 结束列
          r: 0 // 结束范围
        }
      }
    ]
  }
  Object.assign(ws, wchs, merges)
  utils.book_append_sheet(wb, ws, 'Sheet1')
  writeFile(wb, '年假计算表格.xlsx')
  setTimeout(() => {
    loading.value = false
  }, 2000)
}
function handleChange(info: UploadChangeParam) {
  // @ts-ignore
  const files = info.fileList
  if (!files) return
  if (files && files.length) {
    const fileReader = new FileReader()
    fileReader.onload = (ev) => {
      // @ts-ignore
      const data = ev.target.result
      const workbook = read(data, { type: 'binary', cellDates: true, cellText: false, raw: true })
      const worksheet = workbook.Sheets[workbook.SheetNames[0]]
      const table = utils.sheet_to_json(worksheet, { raw: true, header: 1 })
      for (let index = 2; index < table.length; index++) {
        const row = table[index] as Array<any>
        const level = row[2]
        const joinDate = row[3]
        const currentData = row[7]
        const currentDay = currentData ?? dayjs().format('YYYY/MM/DD')
        const workYears = dayjs(currentDay).diff(dayjs(joinDate), 'day') / 365

        // 根据入职的月日，加上当前的年份，得到今年的入职日期
        const currentJoinDate = dayjs()
          .set('month', dayjs(joinDate).month())
          .set('date', dayjs(joinDate).date())

        // 根据joindate当前年自然年到joinDate的天数

        // 年初到入职日期的天数
        const joinDay = currentJoinDate.dayOfYear() - 1

        // 当前入职日期到当前日期的天数
        const remainDay = dayjs(currentDay).isAfter(currentJoinDate)
          ? dayjs(currentDay).diff(dayjs(currentJoinDate), 'day') + 1
          : 0

        // 入职的年份和当前年份是否是同一年
        const isSameYear = dayjs(currentDay).isSame(joinDate, 'year')

        // 年初到当前日期的天数
        let currentYearHoliday = 0
        const currentDayOfYear = dayjs(currentDay).dayOfYear()
        const currentYearRemainDays = 365 - currentDayOfYear
        // 当前年剩余的假期天数
        let currentYearRemainHoliday = 0
        if (level === 'L0') {
          if (workYears < 1) {
            currentYearHoliday = 0
            currentYearRemainHoliday = 0
          } else if (workYears >= 1 && workYears < 2) {
            const holiday = (remainDay * 5) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
          } else if (workYears >= 2 && workYears < 3) {
            currentYearHoliday = (currentDayOfYear * 5) / 365
            currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
          } else if (workYears >= 3 && workYears < 4) {
            const holiday = (joinDay * 5) / 365 + (remainDay * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          } else if (workYears >= 4 && workYears < 10) {
            const holiday = (currentDayOfYear * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          }
        } else if (level === 'L1') {
          if (workYears < 1) {
            if (isSameYear) {
              const holiday = (remainDay * 5) / 365
              currentYearHoliday = round(holiday)
              currentYearRemainHoliday = round((remainDay * 5) / 365)
            } else {
              const holiday = (currentYearRemainDays * 5) / 365
              currentYearHoliday = round(holiday)
              currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
            }
            const holiday = (joinDay * 5) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
          } else if (workYears < 3) {
            const holiday = (currentDayOfYear * 5) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
          } else if (workYears >= 3 && workYears < 4) {
            const holiday = (joinDay * 5) / 365 + (remainDay * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          } else if (workYears >= 4 && workYears < 10) {
            const holiday = (currentDayOfYear * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          }
        } else if (level === 'L2') {
          if (workYears < 1) {
            if (isSameYear) {
              const holiday = (remainDay * 5) / 365
              currentYearHoliday = round(holiday)
              currentYearRemainHoliday = round((remainDay * 5) / 365)
            } else {
              const holiday = (currentYearRemainDays * 5) / 365
              currentYearHoliday = round(holiday)
              currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
            }
          } else if (workYears < 2) {
            const holiday = (currentDayOfYear * 5) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
          } else if (workYears >= 2 && workYears < 3) {
            const holiday = (joinDay * 5) / 365 + (remainDay * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          } else if (workYears >= 3 && workYears < 10) {
            const holiday = (currentDayOfYear * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          }
        } else if (level === 'L3') {
          if (workYears < 1) {
            const holiday = (joinDay * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)

            if (isSameYear) {
              const holiday = (remainDay * 10) / 365
              currentYearHoliday = round(holiday)
              currentYearRemainHoliday = round((remainDay * 10) / 365)
            } else {
              const holiday = (currentYearRemainDays * 10) / 365
              currentYearHoliday = round(holiday)
              currentYearRemainHoliday = round((currentYearRemainDays * 5) / 365)
            }
          } else if (workYears < 10) {
            const holiday = (currentDayOfYear * 10) / 365
            currentYearHoliday = round(holiday)
            currentYearRemainHoliday = round((currentYearRemainDays * 10) / 365)
          }
        }

        if (workYears >= 10 && workYears < 11) {
          const holiday = (joinDay * 10) / 365 + (remainDay * 15) / 365
          currentYearHoliday = round(holiday)
          currentYearRemainHoliday = round((currentYearRemainDays * 15) / 365)
        } else if (workYears >= 11) {
          const holiday = (currentDayOfYear * 15) / 365
          currentYearHoliday = round(holiday)
          currentYearRemainHoliday = round((currentYearRemainDays * 15) / 365)
        }
        row[9] = currentYearHoliday - row[23]
        row[24] = currentYearRemainHoliday

        tableData = table
      }
      tableData = table
      console.log(tableData)
    }
    // @ts-ignore
    fileReader.readAsBinaryString(files[0].originFileObj)
  }
}
</script>

<template>
  <div style="margin-top: -200px">
    <a-upload @change="handleChange" :before-upload="() => false">
      <a-button>
        <upload-outlined></upload-outlined>
        上传年假表格
      </a-button>
    </a-upload>
  </div>
  <div>
    <a-button
      @click="download"
      :loading="loading"
      type="primary"
      style="width: 138px; margin-top: 50px"
      >下载</a-button
    >
  </div>
</template>

<style scoped>
header {
  line-height: 1.5;
}

.logo {
  display: block;
  margin: 0 auto 2rem;
}

@media (min-width: 1024px) {
  header {
    display: flex;
    place-items: center;
    padding-right: calc(var(--section-gap) / 2);
  }

  .logo {
    margin: 0 2rem 0 0;
  }

  header .wrapper {
    display: flex;
    place-items: flex-start;
    flex-wrap: wrap;
  }
}
</style>
