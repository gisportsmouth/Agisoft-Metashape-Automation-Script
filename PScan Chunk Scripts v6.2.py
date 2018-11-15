"""
Created by M.Schaefer
v6.2 31/08/2018
Incorporating USGS Workflow

To do:
    Currently needs to be run as one session, valid ties not accessible after 
    gradual selection. To be fixed in next version using tracks.
    
    len(chunk.point_cloud.points) - valid ties
    len(chunk.point_cloud.tracks) - all ties

    export: if tiff to large process error to log and continue
    
    
With gradual selection:
    After alignment the model is optimised and then a gradual selection is run.
    After the align process all tie points have errors attached
            to them. The values used are a matter of debate. 
            Rec Uncert is normally set to 10 and run only once. If 
            the value of 10 selects more than 50% of the tie points the value 
            self-adjusts to less than 50%.
            Proj acc is set to start at 2 and self adjusts so less 
            than 50% of the tie points left after Rec Uncert are selected.
            For Reprojection Error the aim is to select up to a set percentage 
            (default 80) of the initial tie points. The value used is self 
            selecting. The selection process is in steps of 10% and optimisation
            is run after each step.
            Values used are saved in the log file.
            
Without gradual selection:
    After alignment the model is optimised once.
    
3D Model/Geo: Geo basically is anything where the photos have coordinates or 
    if using Ground Control Points(GCP). If you have photos of objects without 
    that info use 3D Model
    
Set-up workflow:
1. Set up your chunk(s)
    import images
    name chunks (optional - helps in sorting the output)
    save a copy of the PS Doc

Custom workflow:
2. (Geo) Set reference system for each chunk
3. Load script
    chose output folder
    chose file prefix (this gets added to each output file - helps if trying different settings)
4. Under "Change Values" select "Get Current Parameter info" (check console window for output)
    check image quality threshold
    check filtering (options: Mild (default), Moderate, Aggressive, none)
    (Geo) check export CRS (in case you want the output to be different from the CRS of the camera/GCP)

No GCP or checking alignment - runs whole process and exports:
5a. (Geo) Choose "Custom":"Run - Geo" (with or without gradual selection)
or
5b. (3D Model) Choose "Custom":"Run - 3D Model" (with or without gradual selection)

With GCP or scales or if wanting to check alignment:
5. Choose "Custom":"Align only" (with or without gradual selection)
6. Set up GCP for each chunk or set scale
7. Check alignment
8a. (Geo) Choose "Custom":"Run all after alignment - Geo" (with or without gradual selection)
or
8b. (3D Model) Choose "Custom":"Run all after alignment - 3D Model" (with or without gradual selection)

Files generated depend on method used (3D or Geo) and input photos.

To run methods individually as well after opening the script, in console type 
    "ps_doc." and then hit "TAB" to see the options.
"""

import PhotoScan
import os
from datetime import datetime

# Checking compatibility
compatible_major_version = '1.4'
found_major_version = '.'.join(PhotoScan.app.version.split('.')[:2])
if found_major_version != compatible_major_version:
    raise Exception('Incompatible PhotoScan version: {} != {}'
                        .format(found_major_version, compatible_major_version))
if not PhotoScan.app.document.path:
    raise Exception('Document not created yet, please save first')

class PS_Proc(object):
    """
    Create an object of type PhotoScan.app.document.
    Create menu items to process all or part of the Photoscan workflow.
    
    The object methods iterate through the chunks in the document and process 
        each in turn.
    """
    
    def __init__(self, doc, *, 
                    min_qual=.7, 
                    filtering=PhotoScan.MildFiltering,
                    rec_uncert=10,
                    proj_acc = 2,
                    tp_pcnt=.2):
        """
        Initialise the object 
        
        Parameters: The current PhotoScan document
                    The minimum acceptable picture quality
                    The filter method
                    The reconstruction uncertainty
                    The projection accuracy starting point
                    The percentage of points to aim for in gradual selection
        User Input: The output path for the export products
                    The filename prefix for export products
                    The filename will consist of the prefix, the name of 
                        the chunk and the type, (LAS, JPG, TIFF, 
                                                    DEM, Ortho etc.)
                        If the input is JPG the ortho output will be JPG too, 
                        otherwise the ortho will be TIFF.
        
        """
        self.path = PhotoScan.app.getExistingDirectory(
                        'Specify DSM/Ortho/Model export folder:')
        self.prefix = PhotoScan.app.getString(
                        label='Enter file prefix: ', value='')
        self.doc = doc
        self.min_qual = min_qual
        self.filtering = filtering
        self.rec_uncert = rec_uncert
        self.proj_acc = proj_acc
        self.tp_pcnt = tp_pcnt
        self.total_points = {}
        self.exp_crs = 0
        if not len(self.doc.chunks):
                raise Exception("No chunks!")
        self.chunks = self.doc.chunks
        #set logfile name to follow progress
        docname = os.path.basename(str(self.doc).strip("<,>,', Document"))[:-4]
        self.log = os.path.join(self.path, '{}_log_{}.txt'.
                        format(docname,
                                datetime.now().strftime("%Y-%m-%d %H-%M")))
        print('Log file to check progress: {}'.format(self.log))
        
    def info(self):
        """
        Output the current attributes of the object
        """
        print('The current path is: {}'.format(self.path))
        print('The current file name prefix is: {}'.format(self.prefix))
        print('The current minimum acceptable image quality is: {}'
                .format(self.min_qual))
        print('The current filtering method is: {}'.format(self.filtering))
        print('The current projection accuracy is: {}'.format(self.proj_acc))
        print('The current reconstruction accuracy is: {}'
                .format(self.rec_uncert))
        print('The current export CRS (EPSG Number) is [0 = project CRS]: {}'
                .format(self.exp_crs))
        
    def change_pre(self):
        self.prefix = PhotoScan.app.getString(
                        label='Enter file prefix: ', value='')
        
    def run_custom(self):
        """
        Change the object attributes for:
                    The minimum acceptable picture quality
                    The filter method
                    The projection accuracy
                    The reconstruction uncertainty
                    The export CRS
        """
        min_qual = PhotoScan.app.getFloat(
                        label='Enter picture quality threshold', value=0.7)
        filtering = PhotoScan.app.getString(
                        label=('Enter filtering mode (None, Mild, Moderate, '
                        'Aggressive)'), value='Mild')
        if filtering.lower() == 'mild':
            filtering = PhotoScan.MildFiltering
        elif filtering.lower() == 'moderate':
            filtering = PhotoScan.ModerateFiltering
        elif filtering.lower() == 'aggressive':
            filtering = PhotoScan.AggressiveFiltering
        elif filtering.lower() == 'none':
            filtering = PhotoScan.NoFiltering
        rec_uncert = PhotoScan.app.getFloat(
                        label='Enter Reconstruction Uncertainty', value=10) 
        proj_acc = PhotoScan.app.getFloat(
                        label='Enter Projection Accuracy', value=2)
        crs = PhotoScan.app.getInt(
                        label=('Enter export CRS if different from project CRS' 
                                '(EPSG Number)'), value=0)
        #initiate object
        self.min_qual = min_qual
        self.filtering = filtering
        self.proj_acc = proj_acc
        self.rec_uncert = rec_uncert
        self.exp_crs = crs
            
    def get_quality(self):
        """
        Estimate the image quality if not already present
        """
        for _ in self.chunks:
            qual = [i.meta['Image/Quality'] for i in _.cameras]
            if None in qual:
                _.estimateImageQuality(_.cameras)
  
    def disable_bad_pics(self, *, min_qual=None):
        """
        Disable any images below the threshold
        
        Parameter: min_qual=number (optional)
        Dependencies: self.get_quality())
        """
        if not min_qual:
            min_qual = self.min_qual
        for _ in self.chunks:
            self.get_quality()
            selected_cameras = [i for i in _.cameras
                                if float(i.meta['Image/Quality']) < min_qual]
            for camera in selected_cameras:
                camera.selected = True
                camera.enabled = False
            tot_cams = len(_.cameras)
            disabled = len(selected_cameras) 
            print(('{} out of {} disabled due to quality issues'
                    .format(disabled, tot_cams)))
            if disabled > tot_cams * .5:
                PhotoScan.app.messageBox('More than half the images are '
                                            'disabled, check your image quality'
                                            ' settings.')
    
    def iterate_grad(self, chunk, filter, step, 
                        sel_value, threshold):
        """
        Helper function to iterate the gradual selection process
        
        Input: chunk - the chunk to be processed    
                filter - the PhotoScan.PointCloud.Filter object
                step - the value to adjust the selection value by each iter
                sel_value - the value to use in the filter
                threshold - value used to adjust the sel_value if too many 
                    tie points are selected initially
        """
        adjust = 0
        limit_reach = False
        #number of points at start of grad sel
        current_ties = len([i for i in chunk.point_cloud.points])
        hard_limit = self.total_points[str(chunk)] * self.tp_pcnt
        while True:
            filter.selectPoints(sel_value - adjust)
            sel = len([i for i in chunk.point_cloud.points if i.selected])
            pcent_thisrun = 100 * sel / current_ties
            pcent_total = 100 * sel / self.total_points[str(chunk)]
            print('This iter ties selected: {} %, % of total ties selected{} %'.
                        format(pcent_thisrun, pcent_total))
            #if only a few points are selected do not proceed
            if sel < 10:
                limit_reach = True
                adjust = 0
            #if more than threshold ties selected change value
            elif sel > current_ties * threshold / 100:
                adjust += step
                continue
            else:
                #if initial sel_value was acceptable
                if adjust == 0:
                    break
                else:
                    #go back a step if a threshold is reached
                    adjust -= step
                    break
        #check that less than the hard limit set are selected overall
        if current_ties - sel < hard_limit:
                limit_reach = True
                adjust = 0
        #remove selected ties
        if not limit_reach:
            filter.removePoints(sel_value - adjust)
        return(limit_reach, sel_value - adjust)    
                                                
    def grad_sel_preGCP(self, *, rec_uncert=None,
                            proj_acc=None,
                            adapt=True):
        """
        Run through a gradual selection process to remove erroneous tie points
        
        Parameters: rec_uncert=number
                    proj_acc=number
                    (all optional)
        Dependencie: iterate_grad()
        
        Description: After the align process all tie points have errors attached
            to them. The values used are a matter of debate. This 
            method will eliminate tie points based on Reconstruction 
            Uncertainty and Projection Accuracy; Reprojection Error is handled 
            separately. Rec Uncert is normally set to 10 and run only once. If 
            the value of 10 selects more than 50% of the tie points the value 
            self-adjusts to less than 50%.
            Proj acc is set to start at 2 and self adjusts so less 
            than 50% of the tie points left after Rec Uncert are selected.
        """
        if not rec_uncert:
            rec_uncert = self.rec_uncert
        if not proj_acc:
            proj_acc = self.proj_acc
        for _ in self.chunks:
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Running preGCP gradual selection of {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Reconstruction uncertainty: {} \n'\
                                '    Adaptive fitting: {} \n'.
                                            format(rec_uncert,
                                                    adapt)) 
            #create filter
            f = PhotoScan.PointCloud.Filter()
            try:
                total_points = self.total_points[str(_)]
            except KeyError:
                print('A key error occurred in chunk: {}'.format(_))
                #write log information
                with open(self.log, 'a+') as logfile:
                    start_t =  datetime.now()
                    logfile.write('***A key error occurred in chunk: {}***'.
                                    format(_))
                continue
            #1 RecUncert
            print('Gradual selection - Reconstruction Uncertainty')
            f.init(_, criterion=PhotoScan.
                                PointCloud.
                                Filter.
                                ReconstructionUncertainty)
            l_reach, val_rec_uncert = self.iterate_grad(_, f, -1, 
                                                    rec_uncert, 50)
            _.optimizeCameras(adaptive_fitting=adapt)
            #2 ProjAcc
            print('Gradual selection - Projection Accuracy')
            f.init(_, criterion=PhotoScan.PointCloud.Filter.ProjectionAccuracy)
            l_reach, val_proj_acc = self.iterate_grad(_, f, -0.1,
                                                proj_acc, 50)
            _.optimizeCameras(adaptive_fitting=adapt)  
            #write log information
            removed = total_points - len([i for i in _.point_cloud.points])
            pcent = 100 * removed / total_points
            with open(self.log, 'a+') as logfile:
                end_t = datetime.now()
                logfile = open(self.log, 'a+')
                logfile.write('Finished preGCP gradual selection {} at {} \n'.
                                            format(_, 
                                                end_t.strftime('%H:%M:%S')))
                logfile.write('    Reconstruction uncertainty used: {} \n '\
                                '    Projection accuracy value used: {} \n'.
                                            format(val_rec_uncert, 
                                                    val_proj_acc))
                logfile.write('    Final points removed: {}, {} % \n'.
                                                format(removed, pcent))
                logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t))) 
            
    def grad_sel_postGCP(self, *, repro_error=0,
                            adapt=True):
        """
        Run through a gradual selection process to remove erroneous tie points
        
        Parameters: repro_error=number
        Dependencie: iterate_grad()
        Description: After the align process all tie points have errors attached
            to them. The values used are a matter of debate This 
            method will eliminate tie points based on Reprojection Error.
            The aim is to select up to a set percentage(default 80) 
            of the initial tie points. The value used is self selecting. The 
            selection process is in steps of 10% and optimisation is run after 
            each step.
        """
        for _ in self.chunks:
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Running postGCP grad selection of {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Aim to reduce to (%): {} \n'\
                                '    Adaptive fitting (%): {} \n'.
                                            format(self.tp_pcnt * 100,
                                                    adapt))
            f = PhotoScan.PointCloud.Filter()
            try:
                total_points = self.total_points[str(_)]
            except KeyError:
                print('A key error occurred in chunk: {}'.format(_))
                #write log information
                with open(self.log, 'a+') as logfile:
                    start_t =  datetime.now()
                    logfile.write('***A key error occurred in chunk: {}***'.
                                    format(_))
                continue
            #ReproError
            print('Gradual selection - Reprojection Error')
            adjust = 0
            step = 0.02
            f.init(_, criterion=PhotoScan.PointCloud.Filter.ReprojectionError)
            l_reach = False
            val_repro_error = repro_error
            while not l_reach:
                l_reach, val_repro_error = self.iterate_grad(_, f, -0.1,
                                                    val_repro_error, 
                                                    10)
                _.optimizeCameras(adaptive_fitting=adapt)
            #write log information
            removed = total_points - len([i for i in _.point_cloud.points])
            pcent = 100 * removed / total_points
            #write log information
            with open(self.log, 'a+') as logfile:
                end_t = datetime.now()
                logfile = open(self.log, 'a+')
                logfile.write('Finished gradual selection {} at {} \n'.
                                            format(_, 
                                                end_t.strftime('%H:%M:%S')))
                logfile.write('    Reprojection error used: {} \n'.
                                            format(val_repro_error))
                logfile.write('    Starting points: {} \n'\
                                '    Points remaining: {}, \n'\
                                '    Points removed total: {} % \n'.
                                        format(total_points,
                                                len([i for i in 
                                                        _.point_cloud.points]), 
                                                pcent))
                logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t)))
                                                
    def remove_align(self):
        """
        Remove the camera alignment
        """
        for _ in self.chunks:
            for c in _.cameras:
                c.transform = None
 
    def align(self, *, generic=False, 
                        filter = True, 
                        acc = PhotoScan.HighAccuracy,
                        key = 40000,
                        tie = 0,
                        adapt=True):
        #star forces named parameters
        """
        Align images in the document
        
        Parameters: generic=boolean (if not generic, reference pre-selection 
                                    is used)
        """
        if not generic:
            reference = True
        else:
            reference=False
        for _ in self.chunks:
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Aligning {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Reference pre-selection: {} \n'.
                                                format(reference))
                logfile.write('    Accuracy: {} \n'\
                                '    Keypoint limit: {} \n'\
                                '    Tiepoint limit: {} \n'\
                                '    Adaptive fitting: {} \n'.
                                                    format(acc, 
                                                            key, 
                                                            tie, 
                                                            adapt))
            #start matching and aligning
            _.matchPhotos(accuracy=acc, 
                            generic_preselection=generic, 
                            reference_preselection=reference,
                            filter_mask=filter, 
                            keypoint_limit=key, 
                            tiepoint_limit=tie)
            _.alignCameras(adaptive_fitting=adapt)
            _.optimizeCameras(adaptive_fitting=adapt)
            #write log information
            with open(self.log, 'a+') as logfile:
                end_t = datetime.now()
                logfile = open(self.log, 'a+')
                logfile.write('Finished aligning {} at {} \n'.
                                            format(_, 
                                                end_t.strftime('%H:%M:%S')))
                logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t)))
            self.total_points[str(_)] = len([i for i in _.point_cloud.points])
            self.doc.save()
     
    def dense_c(self, *, mode=None, qual=PhotoScan.HighQuality):
        """
        Create  dense point cloud
        
        Parameter: mode=[PhotoScan filtering method]
        """
        if not mode:
            mode = self.filtering
        for _ in self.chunks:
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Building point cloud {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Quality: {} \n'\
                                '    Filtering mode: {} \n'.
                                        format(qual, mode))
            #build depthmaps and dense cloud
            _.buildDepthMaps(quality=qual, 
                                filter=mode)            
            _.buildDenseCloud(point_colors=True)
            #write log information
            with open(self.log, 'a+') as logfile:
                end_t = datetime.now()
                logfile.write('Finished generating point cloud {} at {} \n'.
                                                format(_, 
                                                    end_t.strftime('%H:%M:%S')))
                logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t)))
            self.doc.save()
              
    def build_model(self, *, surf=PhotoScan.Arbitrary,
                                inter=PhotoScan.EnabledInterpolation,
                                face=PhotoScan.MediumFaceCount,
                                map=PhotoScan.GenericMapping,
                                blend=PhotoScan.MosaicBlending,
                                m_size=4096):
        """
        Build a model from the dense point cloud
        """
        for _ in self.chunks:
            #if a MemoryError occurs in chunk other chunks are still processed
            try:
                #write log information
                with open(self.log, 'a+') as logfile:
                    start_t =  datetime.now()
                    logfile.write('Building  Model {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                    logfile.write('    Surface type: {} \n'\
                                    '    Interpolation: {} \n'\
                                    '    Face count: {} \n'\
                                    '    Mapping: {} \n'\
                                    '    Blending: {} \n'\
                                    '    Mosaic size: {} \n'.
                                            format(surf, 
                                                    inter, 
                                                    face, 
                                                    map, 
                                                    blend, 
                                                    m_size))
                #Build model and texture                                
                _.buildModel(surface=surf, 
                                interpolation=inter,
                                face_count=face,
                                source=PhotoScan.DenseCloudData,
                                vertex_colors=True)
                _.buildUV(mapping=map)
                _.buildTexture(blending=blend, size=m_size)
                #write log information
                with open(self.log, 'a+') as logfile:
                    end_t = datetime.now()
                    logfile.write('Finished building model {} at {} \n'.
                                                format(_, 
                                                    end_t.strftime('%H:%M:%S')))
                    logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t)))
                self.doc.save()
            except MemoryError:
                print('A memory error occurred in chunk: {}'.format(_))
                #write log information
                with open(self.log, 'a+') as logfile:
                    start_t =  datetime.now()
                    logfile.write('***A memory error occurred in chunk: {}***'.
                                    format(_))
                continue
            
    def dem(self):
        """
        Build a DEM from the dense point cloud
        """
        for _ in self.chunks:
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Building DEM {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
            #build DEM                            
            _.buildDem(source=PhotoScan.DenseCloudData, 
                        interpolation=PhotoScan.EnabledInterpolation)
            #write log information
            with open(self.log, 'a+') as logfile:
                end_t = datetime.now()
                logfile.write('Finished building DEM {} at {} \n'.
                                                format(_, 
                                                    end_t.strftime('%H:%M:%S')))
                logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t)))
        self.doc.save()        
  
    def ortho(self, *, holes=True):
        """
        Build an ortho from the DEM
        """
        for _ in self.chunks:
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Generating Ortho {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Fill holes = {}'.format(holes))
            _.buildOrthomosaic(surface=PhotoScan.ElevationData, 
                                blending=PhotoScan.MosaicBlending, 
                                fill_holes=holes)
            #write log information
            with open(self.log, 'a+') as logfile:
                end_t = datetime.now()
                logfile.write('Finished generating Ortho{} at {} \n'.
                                                format(_, 
                                                    end_t.strftime('%H:%M:%S')))
                logfile.write('Processing time {} \n'.
                                                format(str(end_t - start_t)))
        self.doc.save()
                                
    def export_geo(self):
        """
        Export LAS, DEM and Ortho to path using the file name prefix
        
        Ortho file format is JPG for JPG input images and TIF for all others, 
            unless the resulting JPG is > 65535 pixels, in which case a TIF will 
            be used.
        """
        jpg_limit = 65535
        for _ in self.chunks:
            ext = _.cameras[0].label[-3:]
            if self.exp_crs == 0:
                crs = _.crs
            else:
                crs = PhotoScan.CoordinateSystem('EPSG::{}'
                        .format(self.exp_crs))
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Exporting geo {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Export CRS: {}\n'.
                                            format(crs))
            #check if input is JPG and below the limit            
            if (ext.upper() == 'JPG' 
                and not (_.orthomosaic.width > jpg_limit 
                        or _.orthomosaic.height > jpg_limit)):
                ext = 'jpg'
            else:
                ext = 'tif'
            if str(_)[8:-4] == 'Chunk':
                name = ''
                num = str(_)[-3]
            else:
                name = str(_)[8:-2]
                num = ''
            try:
                print('las crs: {}'.format(crs))
                file = '{}{}_LAS{}.las'.format(self.prefix, name, num)
                _.exportPoints(path = os.path.join(self.path, file),
                                format=PhotoScan.PointsFormatLAS,
                                projection=crs)
            except RuntimeError as e:
                if str(e) == 'Null point cloud':
                    print('There is no point cloud to export in chunk: {}'
                            .format(_))
                    continue
                else:
                    raise
            try:
                file = '{}{}_DSM{}.tif'.format(self.prefix, name, num)
                _.exportDem(path = os.path.join(self.path, file), 
                            write_world = True,
                            projection=crs)
            except RuntimeError as e:
                if str(e) == 'Null elevation':
                    print('There is no elevation to export in chunk: {}'
                            .format(_))
                    continue
                else:
                    raise
            try:
                file = '{}{}_Ortho{}.{}'.format(self.prefix, name, num, ext)
                _.exportOrthomosaic(path = os.path.join(self.path, file), 
                                    write_world = True,
                                    projection=crs)
            except RuntimeError as e:
                if str(e) == 'Null orthomosaic':
                    print('There is no orthomosaic to export in chunk: {}'
                            .format(_))
                    continue
                else:
                    raise
    
    def export_model(self):
        """
        Export LAS and OBJ to path using the file name prefix
        
        #Texture file format is JPG for JPG input images and TIF for all others.
        """
        for _ in self.chunks:
            #Texture JPG for now as Cloudcompare doesn't like the TIFF format
            ext = PhotoScan.ImageFormatJPEG
            """ext = _.cameras[0].label[-3:]
            if ext.upper() == 'JPG':
                ext = PhotoScan.ImageFormatJPEG
            else:
                ext = PhotoScan.ImageFormatTIFF"""
            if self.exp_crs == 0:
                crs = _.crs
            else:
                crs = PhotoScan.CoordinateSystem('EPSG::{}'
                        .format(self.exp_crs))
            #write log information
            with open(self.log, 'a+') as logfile:
                start_t =  datetime.now()
                logfile.write('Exporting model {} at {} \n'.
                                            format(_, 
                                                start_t.strftime('%H:%M:%S')))
                logfile.write('    Export CRS: {}\n'.
                                            format(crs))
            #create export file name                                
            if str(_)[8:-4] == 'Chunk':
                name = ''
                num = str(_)[-3]
            else:
                name = str(_)[8:-2]
                num = ''
            try:
                file = '{}{}_LAS{}.las'.format(self.prefix, name, num)
                _.exportPoints(path = os.path.join(self.path, file),
                                format=PhotoScan.PointsFormatLAS,
                                projection=crs)
            except RuntimeError as e:
                if str(e) == 'Null point cloud':
                    print('There is no point cloud to export in chunk: {}'
                            .format(_))
                    continue
                else:
                    raise
            try:
                file = '{}{}_OBJ{}.obj'.format(self.prefix, name, num)
                _.exportModel(path = os.path.join(self.path, file),
                                texture_format=ext,
                                projection=crs)
            except RuntimeError as e:
                if str(e) == 'Null model':
                    print('There is no model to export in chunk: {}'.format(_))
                    continue
                else:
                    raise
                 
    def run_geo(self, *, align=True, grad=False):
        """
        Processes georeferenced images, e.g. UAV images
        
        Input:
            align (default True): if True runs the whole process from aligning 
                images to exporting the LAS, Ortho and DSM/Ortho.
                Use Flase if the images are already aligned.
            grad (default False): if True automatically runs a gradual 
                selection and optimisation process
        """
        if align:
            self.disable_bad_pics()
            self.align()
            if grad:
                self.grad_sel_preGCP()
                self.grad_sel_postGCP()
        else:
            if grad:
                self.grad_sel_postGCP()
        self.dense_c()
        self.dem()
        self.ortho()
        self.export_geo()   
               
    def run_model(self, *, align=True, grad=False):
        """
        Processes non-referenced images, e.g. close range object pictures
        
        Input:
            align (default True): if True runs the whole process from aligning 
                images to exporting the LAS and OBJ.
                Use Flase if the images are already aligned.
            grad (default False): if True automatically runs a gradual 
                selection and optimisation process
        """
        if align:
            self.disable_bad_pics()
            self.align()
            if grad:
                self.grad_sel_preGCP()
        else:
            if grad:
                self.grad_sel_postGCP()
        self.dense_c()
        self.build_model()
        self.export_model()  
        
    #different running options for menu    
    def menu_geo_grad(self):
        self.run_geo(grad=True) 
        
    def menu_align_grad(self):
        self.align()
        self.grad_sel_preGCP()
                
    def menu_geo_post_align(self):
        self.run_geo(align=False)
        
    def menu_geo_post_align_grad(self):
        self.run_geo(align=False, grad=True)
  
        
    def menu_model_grad(self):
        self.run_model(grad=True) 
        
    def menu_model_post_align(self):
        self.run_model(align=False)  
        
    def menu_model_post_align_grad(self):
        self.run_model(align=False, grad=True)  
        

def menu(label, method):
    """
    Creates a menu item in PS
    
    Input:  A label for the menu item (string)
            The method in the PS_Proc class to call
    """
    PhotoScan.app.addMenuItem(label, method)
     

#initiate object    
ps_doc = PS_Proc(PhotoScan.app.document)

#create menu options
menu('Custom/Remove alignment(optional)', ps_doc.remove_align)
menu('Custom/Calculate Image Quality', ps_doc.get_quality)
menu('Custom/Align only', ps_doc.align)
menu('Custom/Align only (gradual selection)', ps_doc.menu_align_grad)
menu('Custom/Run - Geo', 
        ps_doc.run_geo)
menu('Custom/Run - Geo (gradual selection)', ps_doc.menu_geo_grad)
menu('Custom/Run all after alignment - Geo', 
        ps_doc.menu_geo_post_align)
menu('Custom/Run all after alignment - Geo (gradual selection)', 
        ps_doc.menu_geo_post_align_grad)
menu('Custom/Run - 3D Model', ps_doc.run_model)
menu('Custom/Run - 3D Model (gradual selection)', ps_doc.menu_model_grad)
menu('Custom/Run all after alignment - 3D Model', ps_doc.menu_model_post_align)
menu('Custom/Run all after alignment - 3D Model (gradual selection)', 
                                            ps_doc.menu_model_post_align_grad)
menu('Change Values/Get current parameter info', ps_doc.info)
menu('Change Values/Change file prefix', ps_doc.change_pre) 
menu('Change Values/Enter custome values', ps_doc.run_custom) 