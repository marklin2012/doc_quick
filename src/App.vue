<script setup>
import PizZip from "pizzip"
import PizZipUtils from "pizzip/utils/index"
import Docxtemplater from "docxtemplater"
import { saveAs } from "file-saver"
import { reactive, ref } from "vue"
import dayjs from "dayjs"
function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback)
}
// const formSize = ref("default")
const formRef = ref()
const contact = reactive({
  guest_chn_name: "",
  guest_eng_name: "",
  guest_passport_number: "",
  address:
    "242/156 Building D, Dcondo Rin, Moo.5, T.Fa Ham,A.Mueang ChiangMai, ChiangMai,50000.",
  rent_start_date: new Date(),
  rent_end_date: new Date(),
  price: "",
  max_guests: 1,
  deposit: "",
  deposit_until: "in 7 days",
})

// const address = computed(() => contact.value)

const rules = reactive({
  guest_chn_name: [
    { required: true, message: "Please input Chinese name", trigger: "blur" },
  ],
  guest_eng_name: [
    { required: true, message: "Please input English name", trigger: "blur" },
  ],
  guest_passport_number: [
    {
      required: true,
      message: "Please input Passport number",
      trigger: "blur",
    },
  ],
  address: [
    { required: true, message: "Please input address", trigger: "blur" },
  ],
  rent_start_date: [
    { required: true, message: "Please input start date", trigger: "blur" },
  ],
  rent_end_date: [
    { required: true, message: "Please input end date", trigger: "blur" },
  ],
  price: [{ required: true, message: "Please input price", trigger: "blur" }],
  max_guests: [
    { required: true, message: "Please input max guest", trigger: "blur" },
  ],
  deposit: [
    { required: true, message: "Please input deposit", trigger: "blur" },
  ],
  deposit_until: [
    { required: true, message: "Please input deposit until", trigger: "blur" },
  ],
})

const submitDoc = async (formEl) => {
  console.log("form", contact)
  await formEl.validate((valid, fields) => {
    if (valid) {
      generateDoc()
    } else {
      console.log("error submit!", fields)
    }
  })
  return
}

const generateDoc = () => {
  loadFile("/RIN156.docx", (error, content) => {
    if (error) {
      console.log(error)
      throw error
    }
    const zip = new PizZip(content)
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    })

    const daysStartDate = dayjs(contact.rent_start_date)
    const daysEndDate = dayjs(contact.rent_end_date)
    const start_year = daysStartDate.format("YYYY")
    const start_month = daysStartDate.format("MM")
    const start_day = daysStartDate.format("DD")
    const start_date = daysStartDate.format("DD/MMM/YYYY")
    console.log("year:", start_year, start_month, start_day, start_date)

    const end_year = daysEndDate.format("YYYY")
    const end_month = daysEndDate.format("MM")
    const end_day = daysEndDate.format("DD")
    const end_date = daysEndDate.format("DD/MMM/YYYY")

    console.log("year:", end_year, end_month, end_day, end_date)

    doc.setData({
      guest_chn_name: contact.guest_chn_name,
      guest_eng_name: contact.guest_eng_name,
      guest_passport_number: contact.guest_passport_number,
      address: contact.address,
      price: contact.price,
      max_guests: contact.max_guests,
      deposit: contact.deposit,
      start_year,
      start_month,
      start_day,
      start_date,
      end_year,
      end_month,
      end_day,
      end_date,
    })
    try {
      // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
      doc.render()
    } catch (error) {
      // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
      function replaceErrors(key, value) {
        if (value instanceof Error) {
          return Object.getOwnPropertyNames(value).reduce(function (
            error,
            key
          ) {
            error[key] = value[key]
            return error
          },
          {})
        }
        return value
      }
      // console.log(JSON.stringify({ error: error }, replaceErrors))

      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map(function (error) {
            return error.properties.explanation
          })
          .join("\n")
        console.log("errorMessages", errorMessages)
        // errorMessages is a humanly readable message looking like this :
        // 'The tag beginning with "foobar" is unopened'
      }
      throw error
    }
    const out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    })
    // Output the document using Data-URI
    saveAs(out, `${contact.guest_chn_name}.docx`)
  })
}
</script>

<template>
  <div class="common-layout">
    <el-container>
      <el-header>
        <el-icon color="#000" :size="30">
          <Document />
        </el-icon>
        <h1>租赁合同生成器</h1>
      </el-header>
      <el-main>
        <el-form ref="formRef" :model="contact" :rules="rules">
          <el-form-item label="Name" prop="guest_chn_name">
            <el-input v-model="contact.guest_chn_name" />
          </el-form-item>
          <el-form-item label="English Name" prop="guest_eng_name">
            <el-input v-model="contact.guest_eng_name" />
          </el-form-item>
          <el-form-item label="Passport No." prop="guest_passport_number">
            <el-input v-model="contact.guest_passport_number" />
          </el-form-item>
          <el-form-item label="Address" prop="address">
            <el-input v-model="contact.address" type="textarea" autosize />
          </el-form-item>
          <el-form-item label="Start date" required>
            <el-date-picker
              v-model="contact.rent_start_date"
              type="date"
              label="Pick a date"
              placeholder="Start date"
              style="width: 100%"
            />
          </el-form-item>
          <el-form-item label="End date" required>
            <el-date-picker
              v-model="contact.rent_end_date"
              type="date"
              label="Pick a date"
              placeholder="End date"
              style="width: 100%"
            />
          </el-form-item>

          <el-form-item label="Price" prop="price">
            <el-input v-model="contact.price" />
          </el-form-item>
          <el-form-item label="Max guests" prop="max_guests">
            <el-input-number v-model="contact.max_guests" :min="1" :max="10" />
          </el-form-item>
          <el-form-item label="Deposit" prop="deposit">
            <el-input v-model="contact.deposit" />
          </el-form-item>

          <el-form-item>
            <el-button type="primary" @click="submitDoc(formRef)">
              Create
            </el-button>
            <el-button @click="resetForm(formRef)">Reset</el-button>
          </el-form-item>
        </el-form>
      </el-main>
    </el-container>
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
