<template>
  <b-container class="bv-example-row">
    <b-row>
      <b-col>
        <b-form @submit="onSubmit">
          <b-form-group
            id="input-group-1"
            label="Excel file"
            label-for="input-1"
          >
            <b-form-file
              v-model="file"
              :state="Boolean(file)"
              placeholder="Choose a file or drop it here..."
              drop-placeholder="Drop file here..."
            ></b-form-file>
          </b-form-group>
          <b-button type="submit" variant="primary">Convert</b-button>
        </b-form>
      </b-col>
    </b-row>
  </b-container>
</template>

<script>
import { read, utils, writeFileXLSX } from "xlsx";

export default {
  data() {
    return {
      file: null,
    };
  },
  methods: {
    onSubmit(event) {
      event.preventDefault();
      const workSheetName = "Okta Users";
      const fr = new FileReader();

      fr.onload = e => {
        const result = JSON.parse(e.target.result);
        const dataList = result.map(item => {
          return {
            firstName: item.profile.firstName,
            lastName: item.profile.lastName,
            email: item.profile.email,
            department: item.profile.department,
            division: item.profile.division,
            role: item.profile.role,
            title: item.profile.title,
          }
        })
  
        const workBook = utils.book_new();
        const workSheetData = dataList
  
        const workSheet = utils.json_to_sheet(workSheetData)
  
        utils.book_append_sheet(workBook, workSheet, workSheetName);
  
        writeFileXLSX(workBook, "export.xlsx");
      }
      fr.readAsText(this.file);
    },
  },
};
</script>
