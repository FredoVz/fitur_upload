<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<title>Fitur Upload Excel to Database</title>
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">
</head>
<body>

<?php if(isset($error)): ?>
    <p style="color: red;"><?php echo $error; ?></p>
<?php endif; ?>

<?php if(isset($success)): ?>
    <p style="color: green;"><?php echo $success; ?></p>
<?php endif; ?>

	<form action="<?php echo base_url('home/index'); ?>" method="post" enctype="multipart/form-data">
		<div class="form-group">
	    	<label>Upload Excel</label>
	    	<input type="file" name="excel" class="form-control">
		</div>
		<button class="btn btn-success">Upload Excel</button>
	</form>

	

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>
</body>
</html>