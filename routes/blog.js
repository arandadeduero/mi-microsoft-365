const express = require('express');
const router = express.Router();
const multer = require('multer');
const { isAuthenticated, checkPermission } = require('../middleware/auth');
const { verifyCsrfToken } = require('../middleware/csrf');
const { auditRequest } = require('../services/auditService');
const { getCategories } = require('../services/wpService');
const { sendEmail } = require('../services/graphService');
const { jetpackEmail } = require('../config/wpConfig');

// Graph /me/sendMail rejects total payloads above 4 MB (after base64 expansion)
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_ATTACHMENT_BYTES, files: 5 },
});

function uploadFiles(req, res, next) {
  upload.array('files', 5)(req, res, (error) => {
    if (!error) return next();
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).send('Cada archivo adjunto puede ocupar como máximo 3 MB.');
    }
    if (error.code === 'LIMIT_FILE_COUNT' || error.code === 'LIMIT_UNEXPECTED_FILE') {
      return res.status(400).send('Puedes adjuntar como máximo 5 archivos.');
    }
    return res.status(400).send('No se han podido procesar los archivos adjuntos.');
  });
}

router.get('/blog', isAuthenticated, checkPermission('/blog'), async (req, res, next) => {
  try {
    const categories = await getCategories();
    res.render('blog', {
      title: 'Publicar en Blog',
      currentPage: 'blog',
      categories
    });
  } catch (error) {
    next(error);
  }
});

router.post('/blog', isAuthenticated, checkPermission('/blog'), uploadFiles, verifyCsrfToken, async (req, res, next) => {
  try {
    const { title, content, categories, delay, status } = req.body;
    const files = req.files;

    // Build email body
    let body = content + '\n\n';
    if (categories) {
      const catArray = Array.isArray(categories) ? categories : [categories];
      body += `[category ${catArray.join(', ')}]\n`;
    }

    body += `[status ${status}]\n`;
    if (delay && delay > 0) {
      body += `[delay +${delay} hours]\n`;
    }

    // Prepare attachments
    const attachments = files.map(file => ({
      name: file.originalname,
      contentType: file.mimetype,
      contentBytes: file.buffer.toString('base64')
    }));

    // Send email
    const emailData = {
      to: jetpackEmail,
      subject: `BANDO: ${title}`,
      body: body,
      attachments: attachments
    };


    await sendEmail(req.session.accessToken, emailData);
    await auditRequest(req, 'blog.post', { title, status, attachments: attachments.length });

    res.redirect('/blog');
  } catch (error) {
    console.error('Error in blog email post:', error);
    res.status(500).send('Error al enviar email para publicar: ' + error.message);
  }
});

module.exports = router;
