# -*- coding: utf-8 -*-
# pycodestyle "H:\My Drive\Python\Metashape Chunk Scripts v8_2.py"
"""
Created by M.Schaefer
v8.3 17/05/2024
v7.8 08/06/2022
v7.7 26/05/2022
v7.6 04/02/2022
v7.5 01/02/2022
v7.4 25/11/2020
v7.3 19/03/2020
v7.2 30/10/2019
v7.1 28/06/2019
v7 13/03/2019
Update to Metashape
Change log:
    v8.4
        Update to major version 2.2
    v8.3
        Update to major version 2.1
    v8.2
        Linted
    v8.1
        Changed default settings:
            Filtering = Moderate
        Added report export
        converted .format to f-string
        replaced os with pathlib.Path
    v8.0
        updated for major version 2.0
        updated script to reflect name changes in 2.0 API
    v7.8
        changed align option for run_model to generic=True
        added parameter to align for guided matching default True
        added align option guided_matching=guided is good for vegetation,
        grass, trees etc., as well as large sensors
        changed align option for run_model to guided=False
        Added blue_flag static method to turn white flags blue
        Added refine seamlines and ghosting =True to ortho process
        Changed export of JPG input to PNG
    v7.7
        Changed default settings:
            Camera quality = 0.5
            Filtering = Aggressive
    v7.6
        Updated to version number 1.8
        Added point_confidence=True to buildDenseCloud
        Added specific workflow and menus for Fjalls, stopping before
            ortho to disable pictures.
    v7.5
        Updated to version number 1.7
    v7.4
        Updated ortho to fix errors due to Metashape python changes
        Updated export_geo to fix errors due to Metashape python changes.
            export-raster projection is now not the metashape crs,
            but an orthoprojection with an attribute called crs
        Added matching accuracy and depth map quality to main class, and a
            custom menu to set these with checks if the user entered values are
            options Metashape supports.
        export: if tiff too large for export try to export as BigTIFF
             and if that fails log error and continue
    v7.3
        Update to 1.6 Python API changes
        For matching accuracy the correspondence should be the following:
        Highest = 0
        High = 1
        Medium = 2
        Low = 4
        Lowest = 8

        For depth maps quality:
        Ultra = 1
        High = 2
        Medium = 4
        Low = 8
        Lowest = 16
    v7.2
    added option to select mask by tie points to model
    added option to run model in highest accuracy and quality
    added resetRegion to end of align process
    v7.1
    added reverse altitude script for DJI "below sea level" issue
    v7
    changed all references to Metashape
    tweaked grad selection process
    made changes to log file - added total time
To do:
    Currently needs to be run as one session, valid ties not accessible after
    gradual selection. To be explored in future version using tracks. But then
    percentages are of all tracks, not the initial valid ones.
        len(chunk.tie_points.points) - valid ties
        len(chunk.tie_points.tracks) - all ties
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
    save a copy of the MS Doc
Custom workflow:
2. (Geo) Set reference system for each chunk
3. Load script
    chose output folder
    chose file prefix (this gets added to each output file -
                       helps if trying different settings)
4. Under "Change Values" select "Get Current Parameter info" (check console
                                                              window for output
                                                              )
    check image quality threshold
    check filtering (options: Mild (default), Moderate, Aggressive, none)
    (Geo) check export CRS (in case you want the output to be different from
    the CRS of the project)
No GCP or checking alignment - runs whole process and exports:
5a. (Geo) Choose "Custom":"Run - Geo" (with or without gradual selection)
or
5b. (3D Model) Choose "Custom":"Run - 3D Model" (with or without gradual
    selection)
With GCP or scales or if wanting to check alignment:
    Having the images aligned helps with marker placement. If markers are
    placed already there is no quality difference as to whether you align on
    image coords first and then switch to GCP, although some sources claim this
    is better. If markers exist you can import them and run the whole process
    to save time (see above).
5. Choose "Custom":"Align only" (with or without gradual selection)
6. Set up GCP for each chunk or set scale
7. Check alignment
8a. (Geo) Choose "Custom":"Run all after alignment - Geo" (with or
                                                    without gradual selection)
or
8b. (3D Model) Choose "Custom":"Run all after alignment - 3D Model" (with or
                                                    without gradual selection)
Files generated depend on method used (3D or Geo) and input photos.
To run methods individually as well after opening the script, in console type
    "ms_doc." and then hit "TAB" to see the options.
"""
# imports
from datetime import datetime
from datetime import timedelta
import time
from pathlib import Path
import Metashape


# custom exceptions
class MSVersionCheck(Exception):
    """Exception if MS is not equal to major version
       supported by this script"""

    def __init__(self, compat, found):
        self.message = ('Incompatible Metashape version: '
                        f'{found} != {compat}')
        super().__init__(self.message)


class MSSaveCheck(Exception):
    """Exception if the MS doc is not saved on disk yet"""

    def __init__(self):
        self.message = 'Metashape document needs to be saved first'
        super().__init__(self.message)


class MSChunckCheck(Exception):
    """Exception if no chunks are found in doc to process"""

    def __init__(self):
        self.message = 'No chunks to process'
        super().__init__(self.message)


# Checking compatibility
COMPATIBLE_MAJOR_VERSION = ['2.1', '2.2']
FOUND_MAJOR_VERSION = '.'.join(Metashape.app.version.split('.')[:2])
if FOUND_MAJOR_VERSION not in COMPATIBLE_MAJOR_VERSION:
    raise MSVersionCheck(COMPATIBLE_MAJOR_VERSION, FOUND_MAJOR_VERSION)
# Check document is saved
if not Metashape.app.document.path:
    raise MSSaveCheck()


class MSProc(object):
    """
    Create an object of type Metashape.app.document.
    Create menu items to process all or part of the Metashape workflow.

    The object methods iterate through the chunks in the document and process
        each in turn.
    """

    def __init__(self,
                 doc,
                 *,
                 min_qual=0.5,
                 filtering=Metashape.ModerateFiltering,
                 rec_uncert=10,
                 proj_acc=2,
                 tp_pcnt=0.2,
                 match_acc=1,
                 depth_qual=2,
                 ):
        """
        Initialise the object

        Parameters: The current Metashape document
                    The minimum acceptable picture quality
                    The filter method
                    The reconstruction uncertainty
                    The projection accuracy starting point
                    The percentage of points to aim for in gradual selection
                    The image matching accuracy
                    The depth map quality
        User Input: The output path for the export products
                    The filename prefix for export products
                    The filename will consist of the prefix, the name of
                        the chunk and the type, (LAS, JPG, TIFF,
                                                    DEM, Ortho etc.)
                        If the input is JPG the ortho output will be PNG,
                        otherwise the ortho will be TIFF.

        """
        self.export_path = Path(Metashape
                                .app
                                .getExistingDirectory('Specify DSM/Ortho/Model'
                                                      ' export folder:'
                                                      )
                                )
        # set output file prefix
        self.prefix = Metashape.app.getString(label='Enter file prefix: ',
                                              value=''
                                              )
        self.doc = doc
        # convert metashape document type to Path
        self.doc_path = Path(str(self.doc)
                             .replace("<Document '", '').replace("'>", '')
                             )
        print(self.doc_path)
        self.min_qual = min_qual
        self.filtering = filtering
        self.rec_uncert = rec_uncert
        self.proj_acc = proj_acc
        self.tp_pcnt = tp_pcnt
        self.match_dict = {0: 'Highest',
                           1: 'High',
                           2: 'Medium',
                           4: 'Low',
                           8: 'Lowest'}
        self.depth_dict = {1: 'Ultra',
                           2: 'High',
                           4: 'Medium',
                           8: 'Low',
                           16: 'Lowest'}
        if match_acc in self.match_dict:
            self.match_acc = match_acc
        else:
            raise ValueError('Unknown Metashape matching accuracy value '
                             '(Highest=0 High=1 Medium=2 Low=4 Lowest=8).'
                             )
        if depth_qual in self.depth_dict:
            self.depth_qual = depth_qual
        else:
            raise ValueError('Unknown Metashape depth map quality value '
                             '(Ultra=1 High=2 Medium=4 Low=8 Lowest=16).'
                             )
        self.total_points = {}
        self.exp_crs = 0
        self.runtime = timedelta(0)
        if not self.doc.chunks:
            raise MSChunckCheck
        self.chunks = self.doc.chunks
        # set logfile name to follow progress
        docname = self.doc_path.stem
        self.log = (self.export_path / f'{docname}'
                    f'_log_{datetime.now().strftime("%Y-%m-%d %H-%M")}.txt'
                    )
        print(f'Log file to check progress: {self.log}')

    def info(self):
        """
        Output the current attributes of the object
        """
        print('Processing menu items')
        print(f'The current export path is: {str(self.export_path)}')
        print(f'The current file name prefix is: {self.prefix}')
        print('The current minimum acceptable image quality is: '
              f'{self.min_qual}'
              )
        print(f'The current filtering method is: {self.filtering}')
        print(f'The current projection accuracy is: {self.proj_acc}')
        print(f'The current reconstruction accuracy is: {self.rec_uncert}')
        print('The current export CRS (EPSG Number) is [0 = project CRS]: '
              f'{self.exp_crs}'
              )
        print('Accuracy menu items')
        print('The current matching accuracy is: '
              f'{self.match_dict[self.match_acc]}'
              )
        print('The current depth map quality is: '
              f'{self.depth_dict[self.depth_qual]}'
              )

    def change_pre(self):
        """Change the export file prefix"""
        self.prefix = Metashape.app.getString(label='Enter file prefix: ',
                                              value=''
                                              )

    def run_custom(self):
        """
        Change the object attributes for:
                    The minimum acceptable picture quality
                    The filter method
                    The projection accuracy
                    The reconstruction uncertainty
                    The export CRS
        """
        min_qual = Metashape.app.getFloat(label='Enter picture quality '
                                                'threshold',
                                          value=0.5
                                          )
        filtering = Metashape.app.getString(label=('Enter filtering mode '
                                                   '(None, Mild, Moderate, '
                                                   'Aggressive)'
                                                   ),
                                            value='Mild'
                                            )
        if filtering.lower() == 'mild':
            filtering = Metashape.MildFiltering
        elif filtering.lower() == 'moderate':
            filtering = Metashape.ModerateFiltering
        elif filtering.lower() == 'aggressive':
            filtering = Metashape.AggressiveFiltering
        elif filtering.lower() == 'none':
            filtering = Metashape.NoFiltering
        rec_uncert = Metashape.app.getFloat(label='Enter Reconstruction '
                                                  'Uncertainty',
                                            value=10
                                            )
        proj_acc = Metashape.app.getFloat(label='Enter Projection Accuracy',
                                          value=2)
        crs = Metashape.app.getInt(label=('Enter export CRS if different '
                                          'from project CRS (EPSG Number)'),
                                   value=0
                                   )
        # initiate object
        self.min_qual = min_qual
        self.filtering = filtering
        self.proj_acc = proj_acc
        self.rec_uncert = rec_uncert
        self.exp_crs = crs

    def run_qual_adjust(self):
        """
        Change the object attributes for:
                    The matching accuracy
                    The depth map quality
        """
        match_acc = Metashape.app.getInt(label='Enter image matching accuracy '
                                               '(Highest=0 High=1 Medium=2 '
                                               'Low=4 Lowest=8)',
                                         value=1,
                                         )
        if match_acc in self.match_dict:
            self.match_acc = match_acc
        else:
            raise ValueError('Unknown Metashape matching accuracy value '
                             '(Highest=0 High=1 Medium=2 Low=4 Lowest=8).'
                             )
        depth_qual = Metashape.app.getInt(label=('Enter depth map quality '
                                                 '(Ultra=1 High=2 Medium=4 '
                                                 'Low=8 Lowest=16)'
                                                 ),
                                          value=2,
                                          )
        if depth_qual in self.depth_dict:
            self.depth_qual = depth_qual
        else:
            raise ValueError('Unknown Metashape depth map quality value '
                             '(Ultra=1 High=2 Medium=4 Low=8 Lowest=16).'
                             )

    def reverse_altitude(self):
        """
        Adds user-defined altitude for camera instances in the Reference pane
        """
        for _ in self.chunks:
            for camera in _.cameras:
                if camera.reference.location:
                    coord = camera.reference.location
                    camera.reference.location = Metashape.Vector(
                        [coord.x, coord.y, coord.z * -1]
                    )

    def get_quality(self):
        """
        Estimate the image quality if not already present
        """
        for _ in self.chunks:
            qual = [i.meta['Image/Quality'] for i in _.cameras]
            if None in qual:
                _.analyzeImages(_.cameras)

    def disable_bad_pics(self,
                         *,
                         min_qual=None
                         ):
        """
        Disable any images below the threshold

        Parameter: min_qual=number (optional)
        Dependencies: self.get_quality())
        """
        if not min_qual:
            min_qual = self.min_qual
        for _ in self.chunks:
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write('Disable pictures (Threshold: '
                              f'{min_qual}) at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
            # calculate image quality
            self.get_quality()
            # find low quality images
            selected_cameras = [i for i in _.cameras
                                if float(i.meta['Image/Quality']) < min_qual
                                ]
            for camera in selected_cameras:
                camera.selected = True
                camera.enabled = False
            # report results
            tot_cams = len(_.cameras)
            disabled = len(selected_cameras)
            print(f'{disabled} out of {tot_cams} '
                  'disabled due to quality issues'
                  )
            # warn if more than half of images selected
            if disabled > tot_cams * 0.5:
                Metashape.app.messageBox('More than half the images are '
                                         'disabled, check your image '
                                         'quality settings.'
                                         )
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write('Finished disabling pictures at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Threshold used: {min_qual} \n ')
                logfile.write(f'    {disabled} out of {tot_cams} '
                              'disabled due to quality issues \n'
                              )
                logfile.write(f'Processing time: {end_t - start_t} / '
                              f'Total Time: {self.runtime} \n'
                              )

    def iterate_grad(self,
                     chunk,
                     ms_filter,
                     step,
                     sel_value,
                     threshold
                     ):
        """
        Helper function to iterate the gradual selection process

        Input: chunk - the chunk to be processed
                ms_filter - the Metashape.TiePoints.Filter object
                step - the value to adjust the selection value by each iter
                sel_value - the value to use in the filter
                threshold - value used to adjust the sel_value if too many
                    tie points are selected initially
        """
        # variable to control the selection process
        adjust = 0
        # report to the calling method that the hard limit is reached
        limit_reach = False
        # number of currenty ties at start of grad sel
        current_ties = len([i for i in chunk.tie_points.points])
        print(f'Starting ties this run: {current_ties}')
        # number of overall ties that need to remain
        hard_limit = self.total_points[str(chunk)] * self.tp_pcnt
        # max selectable current ties
        max_ties = current_ties * threshold / 100
        while True:
            # apply filter
            ms_filter.selectPoints(sel_value - adjust)
            # count selected points
            sel = len([i for i in chunk.tie_points.points if i.selected])
            # calculate remaining ties
            remain_ties = current_ties - sel
            # work out percentages for output
            pcent_thisrun = 100 * sel / current_ties
            print(f'This iter % ties selected: {round(pcent_thisrun, 1)}')
            print('This iter # ties selected/starting ties: '
                  f'{sel}/{current_ties}'
                  )
            # check filter limits
            if remain_ties > hard_limit and sel < max_ties:
                # selected value is ok
                print(f'Acepted value this iter: {sel_value - adjust}')
                break
            elif remain_ties > hard_limit and sel > max_ties:
                # if more than threshold ties selected change value and
                # try again
                adjust += step
                print(
                    f'Adjusted by {round(adjust, 2)} to '
                    f'{round(sel_value - adjust, 2)}'
                )
                continue
            elif remain_ties < hard_limit and sel > max_ties:
                # if more than threshold ties selected change value and
                # try again
                adjust += step
                print(f'Adjusted by {adjust} to {sel_value - adjust}')
                continue
            else:
                # ensure sufficient ties remain
                limit_reach = True
                print('Hard limit reached')
                break
        # remove selected ties
        if not limit_reach:
            ms_filter.removePoints(sel_value - adjust)
        return (limit_reach, sel_value - adjust)

    def grad_sel_pregcp(self,
                        *,
                        rec_uncert=None,
                        proj_acc=None,
                        adapt=True
                        ):
        """
        Run through a gradual selection process to remove erroneous tie points

        Parameters: rec_uncert=number
                    proj_acc=number
                    (all optional)
        Dependencie: iterate_grad()

        Description: After the align process all tie points have errors
            attached to them. The values used are a matter of debate. This
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
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Running preGCP gradual selection of {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Reconstruction uncertainty: '
                              f'{rec_uncert} \n'
                              f'    Adaptive fitting: {adapt} \n'
                              )
            # create filter
            f = Metashape.TiePoints.Filter()
            try:
                total_points = self.total_points[str(_)]
            except KeyError:
                print(f'A key error occurred in chunk: {_}')
                print('If grad_sel is run without the script align step '
                      'then total_points needs to be run manually in console:'
                      'for _ in ms_doc.chunks:'
                      'ms_doc.total_points[str(_)] = '
                      'len([i for i in _.tie_points.points])'
                      )
                # write log information
                with open(self.log, 'a+', encoding='utf-8') as logfile:
                    start_t = datetime.now()
                    logfile.write(f'***A key error occurred in chunk: {_}***')
                continue
            # 1 RecUncert
            print('Gradual selection - Reconstruction Uncertainty')
            f.init(_,
                   criterion=Metashape
                            .TiePoints
                            .Filter
                            .ReconstructionUncertainty
                   )
            l_reach, val_rec_uncert = self.iterate_grad(_,
                                                        f,
                                                        -1,
                                                        rec_uncert,
                                                        50)
            _.optimizeCameras(adaptive_fitting=adapt)
            print('Ties remaining after optimisation: '
                  f'{len([i for i in _.tie_points.points])}'
                  )
            # 2 ProjAcc
            print('Gradual selection - Projection Accuracy')
            f.init(_, criterion=Metashape.TiePoints.Filter.ProjectionAccuracy)
            l_reach, val_proj_acc = self.iterate_grad(_,
                                                      f,
                                                      -0.1,
                                                      proj_acc,
                                                      50)
            _.optimizeCameras(adaptive_fitting=adapt)
            print('Ties remaining after optimisation: '
                  f'{len([i for i in _.tie_points.points])}'
                  )
            # write log information
            removed = total_points - len([i for i in _.tie_points.points])
            pcent = 100 * removed / total_points
            self.doc.save()
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write(f'Finished preGCP gradual selection {_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write('    Reconstruction uncertainty used: '
                              f'{val_rec_uncert} \n '
                              '    Projection accuracy value used: '
                              f'{val_proj_acc} \n'
                              )
                logfile.write('    Final points removed (RU, PA & '
                              f'optimisation): {removed}, '
                              f'{round(pcent,2)} % \n'
                              )
                logfile.write(f'Processing time: {end_t - start_t} / '
                              f'Total Time: {self.runtime} \n'
                              )

    def grad_sel_postgcp(self,
                         *,
                         repro_error=0,
                         adapt=True
                         ):
        """
        Run through a gradual selection process to remove erroneous tie points

        Parameters: repro_error=number
        Dependencie: iterate_grad()
        Description: After the align process all tie points have errors
            attached to them. The values used are a matter of debate. This
            method will eliminate tie points based on Reprojection Error.
            The aim is to select up to a set percentage(default 80)
            of the initial tie points. The value used is self selecting. The
            selection process is in steps of ca. 10% and optimisation is run
            after each step.
        """
        for _ in self.chunks:
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(
                    f'Running postGCP grad selection of {_} at '
                    f'{start_t.strftime("%H:%M:%S")} \n'
                )
                logfile.write(f'    Aim to reduce to (%): '
                              f'{self.tp_pcnt * 100} \n'
                              f'    Adaptive fitting (%): {adapt} \n'
                              )
            # filter
            f = Metashape.TiePoints.Filter()
            try:
                total_points = self.total_points[str(_)]
            except KeyError:
                print(f'A key error occurred in chunk: {_}')
                # write log information
                with open(self.log, 'a+', encoding='utf-8') as logfile:
                    start_t = datetime.now()
                    logfile.write(f'***A key error occurred in chunk: {_}***')
                continue
            # ReproError
            print('Gradual selection - Reprojection Error')
            f.init(_, criterion=Metashape.TiePoints.Filter.ReprojectionError)
            l_reach = False
            val_repro_error = repro_error
            while not l_reach:
                # first run finds the RE that selects ca. 10% that is then
                #    applied in subsequent iterations
                l_reach, val_repro_error = self.iterate_grad(_,
                                                             f,
                                                             -0.01,
                                                             val_repro_error,
                                                             10
                                                             )
                time.sleep(5)
                _.optimizeCameras(adaptive_fitting=adapt)
            # write log information
            removed = total_points - len([i for i in _.tie_points.points])
            pcent = 100 * removed / total_points
            self.doc.save()
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write(f'Finished gradual selection {_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write('    Final Reprojection error used: '
                              f'{round(val_repro_error,2)} \n'
                              )
                logfile.write(f'    Starting points: {total_points} \n'
                              '    Points remaining: '
                              f'{len([i for i in _.tie_points.points])}, \n'
                              '    Points removed total: '
                              f'{round(pcent, 2)} % \n'
                              )
                logfile.write(f'Processing time: {end_t - start_t} / '
                              f'Total Time: {self.runtime} \n'
                              )

    def remove_align(self):
        """
        Remove the camera alignment
        """
        for _ in self.chunks:
            for c in _.cameras:
                c.transform = None

    def align(
        self,
        *,
        generic=False,
        ms_filter=True,
        mask_ties=False,
        acc=None,
        key=40000,
        tie=0,
        adapt=True,
        guided=True,
    ):
        # star forces named parameters
        """
        Align images in the document

        Parameters: generic=boolean (if not generic, reference pre-selection
                                    is used)
        """
        if not generic:
            reference = True
        else:
            reference = False
        if not acc:
            acc = self.match_acc
        if acc not in self.match_dict:
            raise ValueError('Unknown Metashape matching accuracy value '
                             '(Highest=0 High=1 Medium=2 Low=4 Lowest=8).'
                             )
        for _ in self.chunks:
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Aligning {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Reference pre-selection: {reference} \n')
                logfile.write(f'    Accuracy: {acc} \n'
                              f'    Keypoint limit: {key} \n'
                              f'    Tiepoint limit: {tie} \n'
                              f'    Adaptive fitting: {adapt} \n'
                              )
            # start matching and aligning
            _.matchPhotos(downscale=acc,
                          generic_preselection=generic,
                          reference_preselection=reference,
                          filter_mask=ms_filter,
                          mask_tiepoints=mask_ties,
                          guided_matching=guided,
                          keypoint_limit=key,
                          tiepoint_limit=tie,
                          )
            _.alignCameras(adaptive_fitting=adapt)
            _.optimizeCameras(adaptive_fitting=adapt)
            _.resetRegion()
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write(f'Finished aligning {_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(
                    f'Processing time: {end_t - start_t} '
                    f'/ Total Time: {self.runtime} \n'
                )
            self.total_points[str(_)] = len([i for i in _.tie_points.points])
            self.doc.save()

    def dense_c(self, *, mode=None, qual=None):
        """
        Create  dense point cloud

        Parameter: mode=[Metashape filtering method]
                    qual=[Metashape depth map quality]
        """
        if not mode:
            mode = self.filtering
        if not qual:
            qual = self.depth_qual
        if qual not in self.depth_dict:
            raise ValueError('Unknown Metashape depth map quality value '
                             '(Ultra=1 High=2 Medium=4 Low=8 Lowest=16).'
                             )
        for _ in self.chunks:
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Building point cloud {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Quality: {qual} \n'
                              f'    Filtering mode: {mode} \n'
                              )
            # build depthmaps and dense cloud
            _.buildDepthMaps(downscale=qual, filter_mode=mode)
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write('Finished generating depth map {_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              'proceeding to poin cloud generation \n')
            _.buildPointCloud(point_colors=True, point_confidence=True)
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write('Finished generating point cloud {_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'Processing time: {end_t - start_t} / '
                              f'Total Time: {self.runtime} \n'
                              )
            self.doc.save()

    def build_model(self,
                    *,
                    surf=Metashape.Arbitrary,
                    inter=Metashape.EnabledInterpolation,
                    face=Metashape.MediumFaceCount,
                    ms_map=Metashape.GenericMapping,
                    blend=Metashape.MosaicBlending,
                    m_size=4096
                    ):
        """
        Build a model from the dense point cloud
        """
        for _ in self.chunks:
            # if a MemoryError occurs in chunk other chunks are still processed
            try:
                # write log information
                with open(self.log, 'a+', encoding='utf-8') as logfile:
                    start_t = datetime.now()
                    logfile.write(f'Building  Model {_} at '
                                  f'{start_t.strftime("%H:%M:%S")} \n'
                                  )
                    logfile.write(f'    Surface type: {surf} \n'
                                  f'    Interpolation: {inter} \n'
                                  f'    Face count: {face} \n'
                                  f'    Mapping: {ms_map} \n'
                                  f'    Blending: {blend} \n'
                                  f'    Mosaic size: {m_size} \n'
                                  )
                # Build model and texture
                _.buildModel(surface_type=surf,
                             interpolation=inter,
                             face_count=face,
                             source_data=Metashape.DepthMapsData,
                             vertex_colors=True,
                             )
                _.buildUV(mapping_mode=ms_map)
                _.buildTexture(blending_mode=blend, texture_size=m_size)
                # write log information
                with open(self.log, 'a+', encoding='utf-8') as logfile:
                    end_t = datetime.now()
                    self.runtime += end_t - start_t
                    logfile.write(f'Finished building model {_} at '
                                  f'{end_t.strftime("%H:%M:%S")} \n'
                                  )
                    logfile.write(f'Processing time: {end_t - start_t} / '
                                  f'Total Time: {self.runtime} \n'
                                  )
                self.doc.save()
            except MemoryError:
                print(f'A memory error occurred in chunk: {_}')
                # write log information
                with open(self.log, 'a+', encoding='utf-8') as logfile:
                    start_t = datetime.now()
                    logfile.write('***A memory error occurred in '
                                  f'chunk: {_}***')
                continue

    def dem(self):
        """
        Build a DEM from the dense point cloud
        """
        for _ in self.chunks:
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Building DEM {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
            # build DEM
            _.buildDem(source_data=Metashape.PointCloudData,
                       interpolation=Metashape.EnabledInterpolation,
                       )
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write('Finished building DEM {_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'Processing time: {end_t - start_t} / '
                              f'Total Time: {self.runtime} \n'
                              )
        self.doc.save()

    def ortho(self, *, holes=True):
        """
        Build an ortho from the DEM
        """
        for _ in self.chunks:
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Generating Ortho {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Fill holes = {holes}')
            _.buildOrthomosaic(surface_data=Metashape.ElevationData,
                               blending_mode=Metashape.MosaicBlending,
                               fill_holes=holes,
                               refine_seamlines=True,
                               ghosting_filter=True,
                               )
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                end_t = datetime.now()
                self.runtime += end_t - start_t
                logfile.write(f'\n Finished generating Ortho{_} at '
                              f'{end_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'Processing time: {end_t - start_t} / '
                              f'Total Time: {self.runtime} \n'
                              )
        self.doc.save()

    def export_geo(self):
        """
        Export LAS, DEM and Ortho to path using the file name prefix
        Ortho file format is PNG for JPG input images and TIF for all others.
        """
        for _ in self.chunks:
            ext = Path(_.cameras[0].photo.path).suffix
            if self.exp_crs == 0:
                crs = _.crs
            else:
                crs = Metashape.CoordinateSystem(f'EPSG::{self.exp_crs}')
            ortho_proj = Metashape.OrthoProjection()
            ortho_proj.crs = crs
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Exporting geo {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Export CRS: {crs}\n')
                # check if input is JPG and below the limit
                if ext.upper() == '.JPG':
                    ext = '.png'
                else:
                    ext = '.tif'
                if str(_)[8:-4] == 'Chunk':
                    name = ''
                    num = str(_)[-3]
                else:
                    name = str(_)[8:-2]
                    num = ''
                try:
                    # Export DSM
                    file = f'{self.prefix}{name}_DSM{num}.tif'
                    _.exportRaster(path=str(self.export_path / file),
                                   source_data=Metashape.ElevationData,
                                   save_world=True,
                                   projection=ortho_proj,
                                   )
                    t = f'File: {file}\n'
                    print(t)
                    logfile.write(t)
                except RuntimeError as e:
                    if str(e) == 'Null elevation':
                        t = ('ERROR: There is no elevation to export in '
                             f'chunk: {_}\n'
                             )
                        print(t)
                        logfile.write(t)
                    else:
                        raise
                try:
                    # Export Ortho
                    file = f'{self.prefix}{name}_Ortho{num}{ext}'
                    _.exportRaster(path=str(self.export_path / file),
                                   source_data=Metashape.OrthomosaicData,
                                   save_world=True,
                                   projection=ortho_proj,
                                   )
                    t = f'File: {file}\n'
                    print(t)
                    logfile.write(t)
                except RuntimeError as e:
                    if str(e) == 'Null orthomosaic':
                        t = ('ERROR: There is no orthomosaic to '
                             f'export in chunk: {_}\n'
                             )
                        print(t)
                        logfile.write(t)
                    elif str(e).startswith('TIFFWriteTile:'):
                        try:
                            print('Attempting BigTIFF\n')
                            compression = Metashape.ImageCompression()
                            compression.tiff_big = True
                            _.exportRaster(path=str(self.export_path / file),
                                           source_data=Metashape.
                                               OrthomosaicData,
                                           save_world=True,
                                           projection=ortho_proj,
                                           image_compression=compression,
                                           )
                            t = 'File: {file}\n'
                            print(t)
                            logfile.write(t)
                            t = ('WARNING: TIFF is too large, '
                                 'exported as BigTIFF\n'
                                 )
                            print(t)
                            logfile.write(t)
                        except Exception as e2:
                            exception_name = type(e2).__name__
                            t = ('ERROR: TIFF is too large to '
                                 'export as a single chunk, export manually\n'
                                 f'Error code: {exception_name}'
                                 )
                            print(t)
                            logfile.write(t)
                    else:
                        print(f'e:{str(e)}')
                        raise
            # export LAS (and any other mesh/texture) and report
            self.export_model()

    def export_model(self):
        """
        Export LAS and OBJ to path using the file name prefix

        #Texture file format is JPG for JPG input images and TIF for
            all others.
        """
        for _ in self.chunks:
            # Texture JPG for now as Cloudcompare doesn't like the TIFF format
            ext = Metashape.ImageFormatJPEG
            ext = _.cameras[0].label[-3:]
            if ext.upper() == 'JPG':
                ext = Metashape.ImageFormatJPEG
            else:
                ext = Metashape.ImageFormatTIFF
            if self.exp_crs == 0:
                crs = _.crs
            else:
                crs = Metashape.CoordinateSystem(f'EPSG::{self.exp_crs}')
            # write log information
            with open(self.log, 'a+', encoding='utf-8') as logfile:
                start_t = datetime.now()
                logfile.write(f'Exporting LAS & model {_} at '
                              f'{start_t.strftime("%H:%M:%S")} \n'
                              )
                logfile.write(f'    Export CRS: {crs}\n')
                # create export file name
                if str(_)[8:-4] == 'Chunk':
                    name = ''
                    num = str(_)[-3]
                else:
                    name = str(_)[8:-2]
                    num = ''
                try:
                    file = f'{self.prefix}{name}_LAS{num}.las'
                    _.exportPointCloud(path=str(self.export_path / file),
                                       format=Metashape.PointCloudFormatLAS,
                                       crs=crs,
                                       )
                    t = f'File: {file}\n'
                    print(t)
                    logfile.write(t)
                except RuntimeError as e:
                    if str(e) == 'Null point cloud':
                        t = ('There is no point cloud '
                             f'to export in chunk: {_}\n')
                        print(t)
                        logfile.write(t)
                    else:
                        raise
                try:
                    file = f'{self.prefix}{name}_OBJ{num}.obj'
                    _.exportModel(path=str(self.export_path / file),
                                  texture_format=ext, crs=crs
                                  )
                    t = f'File: {file}\n'
                    print(t)
                    logfile.write(t)
                except Exception as e:
                    if str(e) == 'Null model':
                        t = f'There is no model to export in chunk: {_}\n'
                        print(t)
                        logfile.write(t)
                    else:
                        raise
                # export report
                try:
                    file = f'{self.prefix}{name}_Report{num}.pdf'
                    _.exportReport(path=str(self.export_path / file),
                                   title=Path(file).stem,
                                   include_system_info=False,
                                   )
                    t = f'File: {file}\n'
                    print(t)
                    logfile.write(t)
                except Exception as e:
                    print(f'Error exporting report: {e}\n')

    def run_geo(self, *, align=True, grad=False, exp=True):
        """
        Processes georeferenced images, e.g. UAV images

        Input:
            align (default True): if True runs the whole process from aligning
                images to exporting the LAS, Ortho and DSM/Ortho.
                Use align=False if the images are already aligned.
            grad (default False): if True automatically runs a gradual
                selection and optimisation process
            pp_req: if further processing is required set this to True and
                the process will stop before the ortho is generated, e.g.
                for multi-spectral data
        """
        if align:
            self.disable_bad_pics()
            self.align()
            if grad:
                self.grad_sel_pregcp()
                self.grad_sel_postgcp()
        else:
            if grad:
                self.grad_sel_postgcp()
        self.dense_c()
        self.dem()
        self.ortho()
        if exp:
            self.export_geo()

    def ortho_and_exp(self):
        """Export ortho and DSM"""
        self.export_geo()

    def run_fjalls_1(self, *, align=True, grad=True):
        """
        Customised for student workflow on a particular practical. Runs
            process up to DEM.

        Input:
            align (default True): if True runs the whole process from aligning
                images to exporting the LAS, Ortho and DSM/Ortho.
                Use align=False if the images are already aligned.
            grad (default False): if True automatically runs a gradual
                selection and optimisation process
        """
        if align:
            self.disable_bad_pics()
            self.align()
            if grad:
                self.grad_sel_pregcp()
                self.grad_sel_postgcp()
        else:
            if grad:
                self.grad_sel_postgcp()
        self.dense_c()
        self.dem()

    def run_fjalls_2(self):
        """
        Customised for student workflow on a particular practical. Runs
            process from Ortho to export.
        """
        self.ortho()
        self.export_geo()

    def run_model(self,
                  *,
                  align=True,
                  grad=False,
                  mask_ties=False,
                  hi_acc=False,
                  exp=True
                  ):
        """
        Processes non-referenced images, e.g. close range object pictures

        Input:
            align (default True): if True runs the whole process from aligning
                images to exporting the LAS and OBJ.
                Use False if the images are already aligned.
            grad (default False): if True automatically runs a gradual
                selection and optimisation process
        """
        if align:
            self.disable_bad_pics()
            if mask_ties and not hi_acc:
                self.align(mask_ties=True)
            elif mask_ties and hi_acc:
                self.align(mask_ties=True, acc=0)
            elif not mask_ties and hi_acc:
                self.align(acc=0)
            else:
                self.align(generic=True, guided=False)
            if grad:
                self.grad_sel_pregcp()
        else:
            if grad:
                self.grad_sel_postgcp()
        if hi_acc:
            self.dense_c(qual=1)
        else:
            self.dense_c()
        self.build_model()
        if exp:
            self.export_model()

    @staticmethod
    def blue_flag():
        """
        Code from Alexey: https://www.agisoft.com/forum/index.php?topic=14568.0

        Run this static method on the current chunk after setting up enough
        control points for white flags to be mostly correct. This will turn
        them blue to participate in the calculations
        """
        chunk = Metashape.app.document.chunk  # active chunk
        for marker in chunk.markers:
            if not marker.position:
                continue
            point = marker.position
            for camera in [c for c in chunk.cameras
                           if c.transform
                           and c.type
                           == Metashape.Camera.Type.Regular
                           ]:
                if camera in marker.projections.keys():
                    continue  # skip existing projections
                x, y = camera.project(point)
                if (0 <= x < camera.sensor.width) \
                        and (0 <= y < camera.sensor.height):
                    marker.projections[camera] = Metashape \
                                                 .Marker \
                                                 .Projection(Metashape
                                                             .Vector([x, y]),
                                                             False
                                                             )
        print("Script finished")

    # different running options for menu
    def menu_geo_grad(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Geo (gradual selection) \n'
                          f'Started:{start_t} \n')
        self.run_geo(grad=True)

    def menu_geo_grad_noexp(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Geo (gradual selection) No export \n'
                          f'Started:{start_t} \n'
                          )
        self.run_geo(grad=True, exp=False)

    def menu_geo_exp(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Orthophoto and Export \n'
                          f'Started:{start_t} \n'
                          )
        self.ortho_and_exp()

    def menu_fjalls_1(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Fjalls_1 \nStarted:{start_t} \n')
        self.run_fjalls_1(grad=True)

    def menu_fjalls_2(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Fjalls_2 \nStarted:{start_t} \n')
        self.run_fjalls_2()

    def menu_align_only(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Align only \nStarted:{start_t} \n')
        self.disable_bad_pics()
        self.align()

    def menu_align_only_grad(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Align only (gradual selection) \n'
                          f'Started:{start_t} \n'
                          )
        self.disable_bad_pics()
        self.align()
        self.grad_sel_pregcp()

    def menu_geo_post_align(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Run all after alignment - Geo \n'
                          f'Started:{start_t} \n'
                          )
        self.run_geo(align=False)

    def menu_geo_post_align_grad(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Run all after alignment - Geo '
                          '(gradual selection) \n'
                          f'Started:{start_t} \n'
                          )
        self.run_geo(align=False, grad=True)

    def menu_model_grad(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: 3D Model (gradual selection) \n'
                          f'Started:{start_t} \n'
                          )
        self.run_model(grad=True)

    def menu_model_grad_mask(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: 3D Model '
                          '(gradual selection, mask ties) \n'
                          f'Started:{start_t} \n'
                          )
        self.run_model(grad=True, mask_ties=True)

    def menu_model_grad_mask_hi(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: 3D Model (gradual selection, mask ties, '
                          'high accuracy) \n'
                          f'Started:{start_t} \n'
                          )
        self.run_model(grad=True, mask_ties=True, hi_acc=True)

    def menu_model_post_align(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Run all after alignment - 3D Model \n'
                          f'Started:{start_t} \n'
                          )
        self.run_model(align=False)

    def menu_model_post_align_grad(self):
        """MS Menu item"""
        # write log information
        with open(self.log, 'a+', encoding='utf-8') as logfile:
            start_t = datetime.now()
            logfile.write('Menu item: Run all after alignment -'
                          ' 3D Model (gradual selection) \n'
                          f'Started:{start_t} \n'
                          )
        self.run_model(align=False, grad=True)


def menu(label, method):
    """
    Creates a menu item in PS

    Input:  A label for the menu item (string)
            The method in the MSProc class to call
    """
    Metashape.app.addMenuItem(label, method)


# initiate object
ms_doc = MSProc(Metashape.app.document)

# create menu options
menu('Custom/Remove alignment(optional)', ms_doc.remove_align)
menu('Custom/Calculate Image Quality', ms_doc.get_quality)
menu('Custom/Run - Geo', ms_doc.run_geo)
menu('Custom/Run - Geo (gradual selection)', ms_doc.menu_geo_grad)
menu('Custom/Run - Geo (gradual selection) No Export',
     ms_doc.menu_geo_grad_noexp)
menu('Custom/Run - Geo Export after processing', ms_doc.menu_geo_exp)
menu('Custom/Run - 3D Model', ms_doc.run_model)
menu('Custom/Run - 3D Model (gradual selection)', ms_doc.menu_model_grad)
menu('Custom/Run - 3D Model (grad, mask ties)', ms_doc.menu_model_grad_mask)
menu('Custom/Run - 3D Model (grad, mask ties, high accuracy)',
     ms_doc.menu_model_grad_mask_hi,
     )
menu('Custom/Align only', ms_doc.menu_align_only)
menu('Custom/Run all after alignment - Geo', ms_doc.menu_geo_post_align)
menu('Custom/Run all after alignment - 3D Model', ms_doc.menu_model_post_align)
menu('Custom/Align only (gradual selection)', ms_doc.menu_align_only_grad)
menu('Custom/Run all after alignment - Geo (gradual selection)',
     ms_doc.menu_geo_post_align_grad,
     )
menu('Custom/Run all after alignment - 3D Model (gradual selection)',
     ms_doc.menu_model_post_align_grad,
     )
menu('Custom/Run - Fjalls_1', ms_doc.menu_fjalls_1)
menu('Custom/Run - Fjalls_2', ms_doc.menu_fjalls_2)
menu('Custom/Run Blue Flag Function', ms_doc.blue_flag)
menu('Change Values/Get current parameter info', ms_doc.info)
menu('Change Values/Change file prefix', ms_doc.change_pre)
menu('Change Values/Enter custom processing values', ms_doc.run_custom)
menu('Change Values/Enter custom accuracy values', ms_doc.run_qual_adjust)
menu('Change Values/Reverse reference altitude', ms_doc.reverse_altitude)
