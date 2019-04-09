/**
 * @swagger
 * /rfxLineItem/getExcelDetails/{rfqNo}:
 *   get:
 *     tags:
 *       - RFXLineItem
 *     description: Get getExcelDetails
 *     produces:
 *       - application/json
 *     parameters:
 *       - name: rfqNo
 *         description: getExcelDetails object
 *         in: path
 *         required: true
 *         type: number
 *     responses:
 *       200:
 *         description: Successfully created
 */

router.get('/rfxLineItem/getExcelDetails/:rfqNo', RfxLineItemController.getExcelDetails)


module.exports.getExcelDetails = async function (req, res) {
    try {
     //   console.log(req.params.rfqNo,'--------controller-------')
        let result = await RfxLineItemServiceObj.getExcelDetails(req);
        return await res.status(200).json(result);
    } catch (e) {
        return await res.status(400).json({'err' : e});
    }
}


module.exports.getExcelDetails = async function (req) {
    console.log(req.params.rfqNo, '--------controller-------')
    if (req.params.rfqNo) {
        let params = { 'rfqNo': req.params.rfqNo };
      let rfqData = await RfxLineItemObj.getRfxLineItems(params)
     // console.log(rfqData,'test')
        await exceFile.generateExcel(rfqData);
        

    }
    return true;
}
