<?php

class Model_home extends CI_Model {

    public function insert_batch($data) {
		return $this->db->insert_batch('tb_impact_music', $data);
	}

	public function tampil_data() {
		return $this->db->get('tb_download_pusatmusik');
	}
}

?>