import Properties from './DocxTemplaterProperties'
import Studio from 'jsreport-studio'

Studio.addPropertiesComponent(Properties.title, Properties, (entity) => entity.__entitySet === 'templates' && entity.recipe === 'docxtemplater')

Studio.entityEditorComponentKeyResolvers.push((entity) => {
  if (entity.__entitySet === 'templates' && entity.recipe === 'docxtemplater') {
    let officeAsset

    if (entity.docxtemplater != null && entity.docxtemplater.templateAssetShortid != null) {
      officeAsset = Studio.getEntityByShortid(entity.docxtemplater.templateAssetShortid, false)
    }

    return {
      key: 'officeAsset',
      entity: officeAsset,
      props: {
        icon: 'fa-link',
        displayName: `docx asset: ${officeAsset != null ? officeAsset.name : '<none>'}`,
        emptyMessage: 'No docx asset assigned, please add a reference to a docx asset in the properties'
      }
    }
  }
})

Studio.addApiSpec({
  template: {
    docxtemplater: {
      templateAsset: {
        encoding: '...',
        content: '...'
      },
      templateAssetShortid: '...'
    }
  }
})

Studio.previewListeners.push((request, entities) => {
  if (request.template.recipe !== 'docxtemplater') {
    return
  }

  if (Studio.extensions.docxtemplater.options.preview.enabled === false) {
    return
  }

  if (Studio.extensions.docxtemplater.options.preview.showWarning === false) {
    return
  }

  if (Studio.getSettingValueByKey('office-preview-informed', false) === true) {
    return
  }

  Studio.setSetting('office-preview-informed', true)

  Studio.openModal(() => <div>
    We need to upload your docx report to our publicly hosted server to be able to use
    Office Online Service for previewing here in the studio. You can disable it in the configuration, see <a
      href='https://jsreport.net/learn/docxtemplater' target='_blank'>https://jsreport.net/learn/docxtemplater</a> for details.
  </div>)
})
