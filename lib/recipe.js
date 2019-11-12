const Docxtemplater = require('docxtemplater')
const JSZip = require('jszip')
const { response } = require('jsreport-office')
const ImageModule = require('docxtemplater-image-module')
const sizeOf = require('image-size')

module.exports = (reporter, definition) => async (req, res) => {
  if (!req.template.docxtemplater || (!req.template.docxtemplater.templateAsset && !req.template.docxtemplater.templateAssetShortid)) {
    throw reporter.createError(`docxtemplater requires template.docxtemplater.templateAsset or template.docxtemplater.templateAssetShortid to be set`, {
      statusCode: 400
    })
  }

  let templateAsset = req.template.docxtemplater.templateAsset

  if (req.template.docxtemplater.templateAssetShortid) {
    templateAsset = await reporter.documentStore.collection('assets').findOne({ shortid: req.template.docxtemplater.templateAssetShortid }, req)

    if (!templateAsset) {
      throw reporter.createError(`Asset with shortid ${req.template.docxtemplater.templateAssetShortid} was not found`, {
        statusCode: 400
      })
    }
  } else {
    if (!Buffer.isBuffer(templateAsset.content)) {
      templateAsset.content = Buffer.from(templateAsset.content, templateAsset.encoding || 'utf8')
    }
  }
  
  const zip = new JSZip(templateAsset.content)

  /**
   * Creating an instance of docxtemplater supporting inline base64 encoded
   * images only. Size will be determinated by using image-size module.
   */
  const imageModule = new ImageModule({
    getImage: (tagValue) => {
      const base64imageRegex = /^data:image\/(png|jpg|svg|svg\+xml);base64,/
      if (base64imageRegex.test(tagValue) === false) {
        return 'IMAGE NOT SUPPORTED'
      }
      const image = tagValue.replace(base64imageRegex, '')
      return Buffer.from(image, 'base64')
    },
    getSize: (img) => {
      const size = sizeOf(img)
      return [size.width, size.height]
    }
  })

  const docx = new Docxtemplater()
  docx.loadZip(zip)
  docx.attachModule(imageModule)
  docx.setData(req.data)
  docx.render()

  return response({
    previewOptions: definition.options.preview,
    officeDocumentType: 'docx',
    buffer: docx.getZip().generate({ type: 'nodebuffer' })
  }, req, res)
}
